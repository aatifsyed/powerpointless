from io import BytesIO
from typing import BinaryIO

import streamlit as st
from pptx import Presentation as presentation
from pptx.presentation import Presentation

from .core import create_subtitles

st.set_page_config(page_title="Create Subtitles | Powerpointless", page_icon="🍬")

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
    f = st.file_uploader("Text Source")
    if f is not None:
        f: BinaryIO
        try:
            text_source = f.read().decode()
        except Exception as e:
            st.error("Couldn't open file as plain text")
            st.exception(e)
            text_source = None
    else:
        text_source = None
elif source_type == "Input text":
    text_source = st.text_area(label="Write text", value="Hi my name's Imogen") or None

st.write("# Result")
if template is None:
    st.error("Please add a template.")
if text_source is None:
    st.error("Please add a text source.")
if None not in (template, text_source):
    with st.spinner(text="Creating powerpoint..."):
        try:
            generated = create_subtitles(
                template=template, lines=text_source.splitlines()
            )
            b = BytesIO()
            generated.save(b)
            st.download_button(
                label="Generated presentation",
                data=b,
                file_name="generated.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        except Exception as e:
            st.error("Couldn't generate presentation. Please check your inputs")
            st.exception(e)
