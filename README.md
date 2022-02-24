# Powerpointless
A webapp to help my friend convert to and from powerpoint "subtitles" for long scripts.

## Running the subtitle creation webapp
[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/aatifsyed/powerpointless/main/create_subtitles_webapp.py)
```sh
poetry run streamlit run create_subtitles_webapp.py
```
## Running the subtitle extraction webapp
[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/aatifsyed/powerpointless/main/extract_subtitles_webapp.py)
```sh
poetry run streamlit run extract_subtitles_webapp.py
```

## Using the command line tool
```console
$ powerpointless create-subtitles -h        
usage: powerpointless create-subtitles [-h] [-t TEMPLATE] -i INPUT -o OUTPUT

optional arguments:
  -h, --help            show this help message and exit
  -t TEMPLATE, --template TEMPLATE
                        Look at the first slide master from this file. Look at the first provided layout in that master. Create a
                        new slide by populating the first placeholder in that layout. Defaults to an internal layout.
  -i INPUT, --input INPUT
                        A new slide will be created for each line in this file.
  -o OUTPUT, --output OUTPUT
                        Write resulting presentation to this file.

$ powerpointless extract-subtitles -h
usage: powerpointless extract-subtitles [-h] -i INPUT -o OUTPUT

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT, --input INPUT
                        A new line will be created for each textbox in each slide in this powerpoint.
  -o OUTPUT, --output OUTPUT
                        Text file containing the powerpoint contents.
```
