import tkinter as tk
from tkinterdnd2 import DND_FILES
from docx import Document
from collections import Counter
import requests
import os
from PIL import Image
from docx.shared import Inches
import io
import logging
from concurrent.futures import ThreadPoolExecutor

# Set up logging
logging.basicConfig(level=logging.INFO)

# Create a ThreadPoolExecutor for fetching images in parallel
executor = ThreadPoolExecutor(max_workers=5)

def process_file(filepath):
    """
    Function to process the file
    """
    try:
        doc = Document(filepath)
        text = ' '.join([paragraph.text for paragraph in doc.paragraphs])
        keywords = extract_keywords(text)
        for keyword in keywords:
            images = find_images(keyword)
            image_urls = list(executor.map(fetch_image, images))
            for image_url in image_urls:
                insert_image(doc, image_url)
        modified_filepath = 'modified_' + filepath
        doc.save(modified_filepath)
        logging.info(f"Finished processing file: {modified_filepath}")
    except Exception as e:
        logging.error(f"An error occurred while processing the file: {e}")

def extract_keywords(text):
    """
    Function to extract keywords from the text
    """
    words = text.split()
    counter = Counter(words)
    keywords = [word for word, count in counter.items() if count > 5]
    return keywords

def find_images(keyword):
    """
    Function to find images related to the keyword
    """
    url = f"https://commons.wikimedia.org/w/api.php?action=query&format=json&list=search&srsearch={keyword}"
    response = requests.get(url)
    pages = response.json()['query']['search']
    images = []
    for page in pages:
        url = f"https://commons.wikimedia.org/w/api.php?action=query&format=json&prop=images&pageids={page['pageid']}"
        response = requests.get(url)
        page_images = response.json()['query']['pages'][str(page['pageid'])]['images']
        images.extend([image['title'] for image in page_images])
    return images

def fetch_image(image_title):
    """
    Function to fetch the image URL
    """
    url = f"https://commons.wikimedia.org/w/api.php?action=query&format=json&prop=imageinfo&iiprop=url&titles={image_title}"
    response = requests.get(url)
    imageinfo = list(response.json()['query']['pages'].values())[0]['imageinfo'][0]
    image_url = imageinfo['url']
    return image_url

def insert_image(doc, image_url):
    """
    Function to insert the image into the document
    """
    try:
        image_response = requests.get(image_url)
        image_bytes = image_response.content
        with Image.open(io.BytesIO(image_bytes)) as pil_image:
            with pil_image.copy() as resized_image:
                resized_image.thumbnail((Inches(1.25), Inches(1.25)))
                temp_filepath = 'temp.jpg'
                resized_image.save(temp_filepath)
                doc.add_picture(temp_filepath)
        os.remove(temp_filepath)
    except Exception as e:
        logging.error(f"An error occurred while inserting an image: {e}")

def drop(event):
    """
    Function to handle the drop event
    """
    try:
        filepath = event.data
        process_file(filepath)
    except Exception as e:
        logging.error(f"An error occurred during the drop event: {e}")

if __name__ == '__main__':
    root = tk.Tk()
    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', drop)

       label = tk.Label(root, text='Drag and Drop a Word File Here')
    label.pack()

    root.mainloop()



