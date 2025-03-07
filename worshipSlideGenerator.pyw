import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
from moviepy.editor import VideoFileClip
from PIL import Image

def get_first_file(directory, extension):
    """Finds the first file with the given extension in the directory."""
    for file in os.listdir(directory):
        if file.lower().endswith(extension):
            return os.path.join(directory, file)
    return None

def extract_sections(file_path):
    """Reads text from the file and splits it by '$'."""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    return content.split('$')

def find_placeholder_pptx(ppt_folder, section_count):
    """Finds the corresponding placeholder PPTX based on section count."""
    if section_count > 30:
        print("Error: Too many sections. Maximum supported is 31.")
        return None
    filename = f"placeHolder{section_count}.pptx"
    source_path = os.path.join(ppt_folder, filename)
    return source_path if os.path.exists(source_path) else None

def unzip_pptx(pptx_path, extract_dir):
    """Extracts a PowerPoint file into a temporary directory."""
    with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

def edit_slide_text(extract_dir, sections):
    """Replaces placeholders in slide text with actual content."""
    slides_path = os.path.join(extract_dir, "ppt", "slides")
    slide_file = os.path.join(slides_path, "slide1.xml")
    if os.path.exists(slide_file):
        tree = ET.parse(slide_file)
        root = tree.getroot()
        ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
              'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
        
        for i, text in enumerate(sections, start=1):
            for elem in root.findall(".//a:t", namespaces=ns):
                if elem.text and elem.text.strip() == f"PlaceHolder{i}":
                    elem.text = text.strip()
        tree.write(slide_file, encoding='utf-8', xml_declaration=True)

def replace_video_and_thumbnail(extract_dir, video_file):
    """Replaces the existing video and its thumbnail in the extracted PPTX folder."""
    media_folder = os.path.join(extract_dir, "ppt", "media")
    
    if not os.path.exists(media_folder):
        print("Error: Media folder not found in PPTX.")
        return
    
    # Find existing media files
    existing_video = None
    thumbnail_image = None
    for file in os.listdir(media_folder):
        if file.endswith(".mp4"):
            existing_video = file
        elif file.endswith((".png", ".jpg", ".jpeg")):
            thumbnail_image = file
    
    # Replace the video
    if existing_video:
        existing_video_path = os.path.join(media_folder, existing_video)
        shutil.copy(video_file, existing_video_path)
        print(f"Replaced video: {existing_video}")

    # Replace the thumbnail image
    if thumbnail_image:
        thumbnail_path = os.path.join(media_folder, thumbnail_image)
        video_clip = VideoFileClip(video_file)
        first_frame = video_clip.get_frame(0)
        img = Image.fromarray(first_frame)
        img.save(thumbnail_path)
        print(f"Replaced video thumbnail: {thumbnail_image}")

def rezip_pptx(extract_dir, output_pptx):
    """Re-zips the extracted directory back into a PowerPoint file."""
    with zipfile.ZipFile(output_pptx, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for root, _, files in os.walk(extract_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, extract_dir)
                zip_ref.write(file_path, arcname)

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    ppt_folder = os.path.join(script_dir, "pptDB")

    # Get the first .txt and .mp4 file
    lyrics_file = get_first_file(script_dir, ".txt")
    video_file = get_first_file(script_dir, ".mp4")

    if not lyrics_file or not video_file:
        print("Error: Missing .txt or .mp4 file in the directory.")
        return

    # Extract sections from the .txt file
    sections = extract_sections(lyrics_file)
    section_count = len(sections)

    # Find the corresponding placeholder PPTX
    source_pptx = find_placeholder_pptx(ppt_folder, section_count)
    if not source_pptx:
        print("Error: No matching placeholder PPTX found.")
        return

    # Set output filename based on the .txt filename (without extension)
    txt_filename = os.path.splitext(os.path.basename(lyrics_file))[0]
    output_pptx = os.path.join(script_dir, f"{txt_filename}.pptx")

    # Create temporary directory
    temp_dir = os.path.join(script_dir, "temp_ppt")
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    # Process PowerPoint
    shutil.copy(source_pptx, os.path.join(script_dir, "temp.pptx"))
    unzip_pptx(os.path.join(script_dir, "temp.pptx"), temp_dir)
    edit_slide_text(temp_dir, sections)
    replace_video_and_thumbnail(temp_dir, video_file)
    rezip_pptx(temp_dir, output_pptx)

    # Clean up
    shutil.rmtree(temp_dir)
    os.remove(os.path.join(script_dir, "temp.pptx"))

    print(f"Presentation '{txt_filename}.pptx' created successfully!")

if __name__ == "__main__":
    main()
