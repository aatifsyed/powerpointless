import streamlit as st
from pptx import Presentation as presentation
from pptx.presentation import Presentation

st.set_page_config(page_title="Powerpointless", page_icon="🍬")

st.write(
    "# Template\n"
    "First add a powerpoint file to use as a template. "
    "Powerpoints may have multiple _Master slide_ sets. "
    "The first such set will be used. "
    "In that set, the first _layout_ (typically the Title Slide layout) will be used. "
    "This will be duplicated for each slide. ",
)

template = st.file_uploader(
    label="Template",
    type=["pptx"],
    accept_multiple_files=False,
)

if template is not None:
    try:
        template: Presentation = presentation(template)
    except Exception as e:
        st.error("Couldn't open template presentation.")
        st.exception(e)
        template = None

st.write(
    "# Text Source\n"
    "Now add the source text. "
    "This must be a plain text file (or you may input text!). "
    "For each line in the text file a new slide will be created. "
    "That template slide will have its first placeholder (text box) populated with the contents of the line. "
)

source_type = st.radio(
    label="Select source type", options=["Upload file", "Input text"]
)
if source_type == "Upload file":
    text_source = st.file_uploader("Text Source")
elif source_type == "Input text":
    text_source = st.text_area(label="Write text", value="Hi my name's Imogen")
