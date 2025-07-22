import pytesseract
import cv2
import pyttsx3
from spire.presentation.common import *
from spire.presentation import *
import os
import PIL.Image
import google.generativeai as genai
import base64
from io import BytesIO
from PIL import Image
import warnings
from gtts import gTTS
warnings.filterwarnings('ignore')

genai.configure(api_key="AIzaSyD_1wJt_juTT2KlHzt6TT0AM87JX2JKpos")   #generate a new key
model = genai.GenerativeModel('gemini-1.5-flash')
pytesseract.pytesseract.tesseract_cmd=r'C:\Program Files\Tesseract-OCR\tesseract.exe' #not required in ubuntu
import streamlit as st
# Function to convert image to base64
def image_to_base64(image):
    buffered = BytesIO()
    image.save(buffered, format="PNG")
    return base64.b64encode(buffered.getvalue()).decode()
image = Image.open(r"D:\Blind_student_presentation\6e1edff7-6c58-497a-a665-735298bcf7fd.jpg")
base64_image = image_to_base64(image)
background_css = f"""
<style>
    .stApp {{
        background-image: url("data:image/png;base64,{base64_image}");
        
        background-position: center;
        background-repeat: no-repeat;
        background-attachment: fixed;
        background-size: cover;
       
}}
</style>
    """
st.markdown(background_css, unsafe_allow_html=True)

custom_title = "Presentation Voice Descriptor using AI"
title_color = "black"
title_size = "40px"

st.markdown(f"<h1 style='color: {title_color}; font-size: {title_size};'>{custom_title}</h1>", unsafe_allow_html=True)
# multi = '''...
# '''
# st.markdown(multi)

# Upload the PowerPoint file
uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

if uploaded_file:
    # Load the PowerPoint file
    presentation = Presentation(uploaded_file)
    st.success("PowerPoint file uploaded successfully!")
    #st.write(uploaded_file.name)
    # Display content slide by slide
    # for slide_num, slide in enumerate(presentation.Slides, start=1):
    #     st.subheader(f"Slide {slide_num}")
        
    #     # Extract and display text
    #     slide_text = []
    #     for shape in slide.shapes:
    #         if shape.has_text_frame:
    #             slide_text.append(shape.text)
    #     st.write("\n".join(slide_text))

# # # Create a Presentation object
    presentation = Presentation()
# # # Load a PowerPoint presentation
    presentation.LoadFromFile(uploaded_file.name)
# Loop through the slides in the presentation
    os.makedirs('D:\Blind_student_presentation\Output',exist_ok=True)
    for i, slide in enumerate(presentation.Slides):
     # Specify the output file name
     fileName ="Output/ToImage_"+ str(i)+ ".png"
     lst=[]
     lst.append(fileName)
     
#     # Save each slide as a PNG image
     image = slide.SaveAsImage()
     image.Save(fileName)
     image.Dispose()
    presentation.Dispose()
    #st.write(lst)
    obj=pyttsx3.init()


    for file in os.listdir("Output"):
      print(file)
   
      extracted=pytesseract.image_to_string('D:\Blind_student_presentation\Output'+'/'+file)
      #st.write(extracted)
      txt_sp=pyttsx3.init()
      if 'Diagram' not in extracted:
                # Convert text to speech
        tts = gTTS(extracted, lang='en')
        
        # Save to a BytesIO object
        audio_file = BytesIO()
        tts.write_to_fp(audio_file)
        audio_file.seek(0)
        
        # Display the audio player
        st.audio(audio_file, format="audio/mp3")
        
        txt_sp.say(extracted)
        txt_sp.runAndWait()
        
      else:
        img=PIL.Image.open('D:\Blind_student_presentation\Output'+'/'+file)
        content="Generate image description"
        response = model.generate_content([content,img])
        tts = gTTS(response.text, lang='en')
        
        # Save to a BytesIO object
        audio_file = BytesIO()
        tts.write_to_fp(audio_file)
        audio_file.seek(0)   
        
        # Display the audio player
        st.audio(audio_file, format="audio/mp3")
        #print(response.text)
        txt_sp.say(response.text)
        txt_sp.runAndWait()
        