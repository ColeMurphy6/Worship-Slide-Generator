# Worship-Slide-Generator
Generates a PowerPoint slide with a video background and overlayed lyrics using animation to reveal each one at a time.

Run worshipSlideGenerator.pyw with a video (ending in ".mp4") and a text file with lyrics (ending in ".txt") in the same directory. Make sure the "pptDB" folder is a folder in the same directory as the script as well. It will extract the lyrics from the text file, split them using the delimeter "$", and make a copy of a slideshow with the same number of text boxes as there are lyric sections from pptDB, changing out the placeholder text and video.

Uses moviepy 1.0.3. If you don't want to worry about dependencies, use the .bat file that creates a python environment.