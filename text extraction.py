import cv2
import easyocr
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
from google.colab import files
import ipywidgets as widgets
from IPython.display import display, clear_output

# Initialize EasyOCR Reader
reader = easyocr.Reader(['en'], gpu=False)  # Force CPU mode in Google Colab

# Function to display image with bounding boxes
def display_image_with_boxes(image_path, results):
    img = cv2.imread(image_path)

    # Draw bounding boxes around detected words
    for (bbox, text, prob) in results:
        (top_left, top_right, bottom_right, bottom_left) = bbox
        top_left = tuple(map(int, top_left))
        bottom_right = tuple(map(int, bottom_right))

        # Draw rectangle and label
        cv2.rectangle(img, top_left, bottom_right, (0, 255, 0), 2)
        cv2.putText(img, text, top_left, cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 0, 0), 2)

    # Convert BGR to RGB (OpenCV uses BGR by default)
    img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

    # Display the image with Matplotlib
    plt.figure(figsize=(10, 6))
    plt.imshow(img_rgb)
    plt.axis("off")  # Hide axes
    plt.show()

# Function to extract text from image using EasyOCR
def extract_text(image_path):
    results = reader.readtext(image_path)
    return results

# Function to handle file upload and text extraction
def upload_and_extract():
    uploaded = files.upload()  # Prompts user to upload an image
    global image_path
    image_path = list(uploaded.keys())[0]  # Get the uploaded file path

    # Extract text from the image
    global results
    results = extract_text(image_path)

    # Display image with bounding boxes around detected words
    display_image_with_boxes(image_path, results)

    # Show detected words with checkboxes for selection
    display_word_selection()

# Function to display interactive checkboxes for word selection
def display_word_selection():
    global checkboxes, submit_button, output_area

    # Clear previous outputs
    clear_output(wait=True)

    # Display the image again
    display_image_with_boxes(image_path, results)

    # Create checkboxes for each detected word
    checkboxes = [widgets.Checkbox(value=False, description=text[1]) for text in results]

    # Arrange checkboxes in two columns
    grid = widgets.GridBox(checkboxes, layout=widgets.Layout(grid_template_columns="repeat(2, 50%)"))

    # Create submit button
    submit_button = widgets.Button(description="Extract Selected Words", button_style='success')

    # Create output area
    output_area = widgets.Output()

    # Display UI
    display(grid, submit_button, output_area)

    # Add event listener for the button
    submit_button.on_click(process_selection)

# Function to process the selected words
def process_selection(b):
    selected_words = [cb.description for cb in checkboxes if cb.value]

    with output_area:
        clear_output(wait=True)
        if selected_words:
            print("\n✅ Selected Words:\n", " ".join(selected_words))
            save_text_to_file(" ".join(selected_words))
        else:
            print("\n⚠ No words selected. Please select at least one word.")

# Function to save and download selected words
def save_text_to_file(text):
    file_name = "selected_words.txt"
    with open(file_name, "w") as file:
        file.write(text)

    # Download file
    files.download(file_name)

# Run the function to upload image and extract words
upload_and_extract()