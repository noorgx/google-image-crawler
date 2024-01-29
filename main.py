import sys
import pandas as pd
from deep_translator import GoogleTranslator
import re
from PIL import Image
import requests
from io import BytesIO
from tqdm import tqdm


def check_color_for_images_parallel(image_info_list, rgb_to_color_func):
    color_info_list = []

    for image_info in image_info_list:
        try:
            # Open the image using Pillow
            image = Image.open(BytesIO(requests.get(image_info['url']).content))

            # Get the resolution (width x height) of the image
            resolution = image.size

            # Get the center coordinates
            center_x = image.width // 2
            center_y = image.height // 2

            # Get the pixel colors at the center
            center_color = image.getpixel((center_x, center_y))

            # Map center color to a named color using rgb_to_color_func
            color_name = rgb_to_color_func(center_color)

            # Append color information to the list
            color_info_list.append({
                'url': image_info['url'],
                'resolution': resolution,
                'center_color': center_color,
                'color_name': color_name
            })


        except Exception as e:
            print(f"Error while checking color: {e}")

    return color_info_list


def sort_by_resolution(image_list):
    # Define a custom key function to extract the resolution values
    def resolution_key(image_dict):
        return tuple(map(int, image_dict['resolution']))

    # Sort the list of dictionaries based on the resolution key in reverse order
    sorted_images = sorted(image_list, key=resolution_key, reverse=True)

    return sorted_images


def sort_and_filter_images(color_info_list, specific_color=None):
    # color, is_specific = specific_color
    # specified_colors = {'gray', 'pink', 'purple'}
    # # Sort the color_info_list based on resolution
    sorted_color_info_list = sorted(color_info_list, key=lambda x: x['resolution'], reverse=True)
    # Filter images with a specific color if specified
    # if is_specific:
    #     filtered_color_info_list = [color_info for color_info in sorted_color_info_list if
    #                                 color_info['color_name'] == color or (
    #                                             color == 'purple' and color_info['color_name'] in specified_colors)]
    # else:
    #     filtered_color_info_list = sorted_color_info_list

    return sorted_color_info_list


def rgb_to_color_name(rgb):
    color_mapping = {
        'red': (255, 0, 0),
        'orange': (255, 165, 0),
        'yellow': (255, 255, 0),
        'green': (0, 255, 0),
        'blue': (0, 0, 255),
        'purple': (128, 0, 128),
        'pink': (255, 182, 193),
        'brown': (165, 42, 42),
        'black': (0, 0, 0),
        'white': (255, 255, 255),
        'gray': (128, 128, 128),
        'grey': (128, 128, 128),
    }

    min_distance = float('inf')
    closest_color = None

    for color, target_rgb in color_mapping.items():
        distance = sum((a - b) ** 2 for a, b in zip(rgb, target_rgb)) ** 0.5
        if distance < min_distance:
            min_distance = distance
            closest_color = color

    return closest_color


def find_and_print_data_tbnid(response_text):
    pattern = r'<div[^>]*data-tbnid="([^"]*)"[^>]*>'
    matches = re.findall(pattern, response_text)
    return matches


def check_for_phrase(text):
    pattern = re.compile(r'Looks like there aren’t any matches for your search', re.IGNORECASE)
    match = re.search(pattern, text)

    if match:
        return True
    else:
        return False


def translate_and_find_color_name(text):
    valid_color_names = {'red', 'orange', 'yellow', 'green', 'blue', 'purple', 'pink', 'brown', 'black', 'white',
                         'gray', 'grey'}
    translated_text = GoogleTranslator(source='auto', target='en').translate(text)
    words = re.findall(r'\b\w+\b', translated_text.lower())

    for word in words:
        if word in valid_color_names:
            return word, True

    return None, False


def google_image_search_0(query, info_color):
    color, specific_color = info_color
    url = "https://www.google.com/search"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/91.0.4472.124 Safari/537.36",
    }

    params = {
        "tbm": "isch",
        "as_q": query + f" -intitle:\"logo\"",
        "tbs": f"ic:specific,isc:{color}" if specific_color else "",
    }

    response = requests.get(url, headers=headers, params=params)

    # Check if the request was successful (status code 200 OK)
    if response.status_code == 200:
        return response.text
    else:
        return None


def google_image_search_1(query, info_color):
    color, specific_color = info_color
    url = "https://www.google.com/search"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/91.0.4472.124 Safari/537.36",
    }

    params = {
        "tbm": "isch",
        "as_q": query + f" intitle:\"{color}\"" if specific_color else query,
        "tbs": f"ic:specific,isc:{color}" if specific_color else "",
    }

    response = requests.get(url, headers=headers, params=params)

    # Check if the request was successful (status code 200 OK)
    if response.status_code == 200:
        return response.text
    else:
        return None


def extract_links_from_page(page_content, any_tag):
    # Define the regex pattern to find script tags with the specified text
    script_pattern = re.compile(fr'<script.*?{any_tag[0] if any_tag else ""}.*?</script>', re.DOTALL)

    # Find all matches of the pattern in the page content
    script_matches = script_pattern.findall(page_content)

    # Define the regex pattern to extract links with the specified sequence
    link_pattern = re.compile(r'(\["https?://[^\s]+",\d+,\d+\])')
    links = []
    # Extract links with the specified sequence from each script tag
    for script_match in script_matches:
        link_matches = link_pattern.findall(script_match)
        for link_match in link_matches:
            second_links = re.findall(r'"(https?://[^"]+)"', link_match)
            pattern = re.compile(r'(\d+),(\d+)')
            matches = pattern.findall(link_match)
            if len(second_links) == 2:
                links.append({
                    'url': second_links[1],
                    'resolution': matches[1],
                })
    return links


def extract_first_three_cells(input_excel, output_excel):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_excel)

    # Use tqdm to create a progress bar for rows
    for index, row in tqdm(df.iterrows(), total=len(df), desc="Processing rows"):
        names = row[:3].tolist()
        i = 1
        image_urls = []
        for name in names:
            if i == 1:
                page_content = google_image_search_0(name, translate_and_find_color_name(names[1]))
                any_tag = find_and_print_data_tbnid(page_content)
                image_urls += [url for url in extract_links_from_page(page_content, any_tag)]
                i += 1
            else:
                page_content = google_image_search_1(name, translate_and_find_color_name(name))
                any_tag = find_and_print_data_tbnid(page_content)
                image_urls += [url for url in extract_links_from_page(page_content, any_tag)]

        # color_info_list = check_color_for_images_parallel(image_urls, rgb_to_color_name)
        # sorted_list = sort_and_filter_images(color_info_list, translate_and_find_color_name(names[1]))
        sorted_list = sort_by_resolution(image_urls)
        image_urls = [url['url'] for url in sorted_list]
        # Check if image_urls is empty before assignment
        if image_urls:
            # Overwrite the content in the 3rd to 7th columns
            df.iloc[index, 3:3 + (5 if len(image_urls) >= 5 else len(image_urls))] = [str(url) for url in
                                                                                      image_urls[
                                                                                      :(5 if len(
                                                                                          image_urls) >= 5 else len(
                                                                                          image_urls))]]  # check for len of the image array if it's less than 5

            # Check if the column exists, create it if it doesn't
            if 'remaining_urls' not in df.columns:
                df['remaining_urls'] = ''

            # Check if the column exists and the index is within the bounds
            if 'remaining_urls' in df.columns and index < len(df):
                if len(image_urls) > 5:
                    df.at[index, 'remaining_urls'] = '\n'.join(
                        str(url) for url in image_urls[5:])  # add the remaining urls
                df.to_excel(output_excel, index=False)
            else:
                # print("Index or column does not exist. Skipping row:", index)
                ""
        else:
            # print("No image URLs found for row:", index)
            ""


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python script.py input_excel output_excel_name.xlsx")
    else:
        input_excel_file = sys.argv[1]
        output_excel_file = sys.argv[2]
        extract_first_three_cells(input_excel_file, output_excel_file)
