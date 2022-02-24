from io import BytesIO

import streamlit as st
from pptx import Presentation as presentation
from pptx.presentation import Presentation

from powerpointless.core import extract_subtitles

st.set_page_config(page_title="Extract Subtitles | Powerpointless", page_icon="🍬")

st.write(
    "# Powerpoint Source\n"
    "A new line will be written for each shape which contains text in this powerpoint. "
    "Note that powerpoint will represent new lines within a shape as a `vertical tab`. "
    "If the shape's text has a trailing whitespace (including newlines and tabs), it will be stripped. ",
)

source = st.file_uploader(
    label="Powerpoint Source",
    type=["pptx"],
    accept_multiple_files=False,
)

if source is not None:
    try:
        source: Presentation = presentation(source)
    except Exception as e:
        st.error("Couldn't open source presentation.")
        st.exception(e)
        source = None

st.write("# Extracted Subtitles\n")

if source is None:
    st.error("Please add a source powerpoint")
else:
    result = "\n".join(extract_subtitles(source))
    result = st.text_area(label="Extracted subtitles", value=result)
    b = BytesIO()
    b.write(result.encode("utf-8"))
    st.download_button(
        label="Download",
        data=b,
        file_name="extracted.txt",
        mime="text/plain",
    )
