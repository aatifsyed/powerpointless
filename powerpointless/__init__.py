import argparse
import logging
import sys
from typing import BinaryIO, List, Optional, TextIO, IO, Any

import argcomplete
from pptx import Presentation as presentation
from pptx.presentation import Presentation
from pptx.slide import (
    Slide,
    SlideLayout,
    SlideMaster,
    SlidePlaceholders,
    Slides,
    SlideShapes,
)


def build_slides(template: Presentation, lines: List[str]) -> Presentation:
    try:
        slide_master: SlideMaster = template.slide_master
    except Exception as e:
        raise RuntimeError(
            "Provided presentation must have at least one slide master attached"
        ) from e

    try:
        layout: SlideLayout = slide_master.slide_layouts[0]
    except Exception as e:
        raise RuntimeError("Slide master must have at least one layout") from e

    slides: Slides = template.slides

    for line in lines:
        new_slide: Slide = slides.add_slide(layout)
        shapes: SlideShapes = new_slide.shapes
        placeholders: SlidePlaceholders = shapes.placeholders
        try:
            placeholders[0].text = line
        except Exception as e:
            raise RuntimeError(
                "Provided layout must use a textbox as the first placeholder"
            )
    return template


class MyFileType(argparse.FileType):
    def __call__(self, string: str) -> IO[Any]:
        """stdlib argparse doesn't handle binary mode..."""
        if string == "-":
            if "r" in self._mode:
                if "b" in self._mode:
                    return sys.stdin.buffer
                else:
                    return sys.stdin
            elif "w" in self._mode:
                if "b" in self._mode:
                    return sys.stdout.buffer
                else:
                    return sys.stdout
        else:
            return super().__call__(string)


logger = logging.getLogger(__name__)


def cli_main() -> int:
    logger.addHandler(logging.StreamHandler())
    logger.setLevel("DEBUG")

    parser = argparse.ArgumentParser(
        description="Creates a slide per line of a user provided file. "
    )
    parser.add_argument(
        "-t",
        "--template",
        type=argparse.FileType(mode="rb"),
        help="Look at the first slide master from this file. "
        "Look at the first provided layout in that master. "
        "Create a new slide by populating the first placeholder in that layout. "
        "Defaults to an internal layout. ",
        default=None,
    )
    parser.add_argument(
        "-i",
        "--input",
        type=MyFileType("r"),
        help="A new slide will be created for each line in this file. ",
        required=True,
    )
    parser.add_argument(
        "-o",
        "--output",
        type=MyFileType("wb"),
        help="Write resulting presentation to this file. ",
        required=True,
    )
    argcomplete.autocomplete(parser)
    args = parser.parse_args()

    logger.debug(args)

    template: Optional[BinaryIO] = args.template
    input: TextIO = args.input
    output: BinaryIO = args.output

    try:
        tmpl: Presentation = presentation(template)
    except Exception as e:
        raise RuntimeError("Couldn't open file as a powerpoint") from e

    prs = build_slides(template=tmpl, lines=input.readlines())

    prs.save(output)


def _main():
    sys.exit(cli_main())
