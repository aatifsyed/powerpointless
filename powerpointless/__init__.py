import argparse
import logging
import sys
from typing import IO, Any, BinaryIO, List, Optional, TextIO

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

logger = logging.getLogger(__name__)


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


def common_main(input: TextIO, output: BinaryIO, template: Optional[BinaryIO]):
    try:
        tmpl: Presentation = presentation(template)
    except Exception as e:
        raise RuntimeError("Couldn't open file as a powerpoint") from e

    prs = build_slides(template=tmpl, lines=input.readlines())

    prs.save(output)


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
        return super().__call__(string)


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

    common_main(input=args.input, output=args.output, template=args.template)

    return 0


def gui_main() -> int:
    logger.addHandler(logging.StreamHandler())
    logger.setLevel("DEBUG")
    import traceback
    from tkinter import Tk, filedialog, messagebox

    root = Tk()
    root.overrideredirect(True)  # Stop flickering
    root.withdraw()  # Hide

    try:
        if messagebox.askyesno(
            title="Do you wish to provide a template?",
            message="Do you wish to provide a template?\n\n"
            "The template must be a pptx file. "
            "The first master slide's first template slide's first placeholder will be populated for each new slide.",
            parent=root,
        ):
            template: Optional[BinaryIO] = filedialog.askopenfile(
                mode="rb", title="Select template powerpoint.", parent=root
            )
        else:
            template = None

        input: Optional[TextIO] = filedialog.askopenfile(
            mode="r",
            title="Select input file. " "A slide will be created per line. ",
            parent=root,
        )

        output: Optional[BinaryIO] = filedialog.asksaveasfile(
            mode="wb",
            confirmoverwrite=True,
            title="Select destination file. "
            "This will contain the final slides, according to the template. ",
            parent=root,
        )
        logger.debug(f"{template=}, {input=}, {output=}")

        if input is None or output is None:
            messagebox.showerror(
                "Input or output not specified. ",
                message="Input or output not specified. " f"{input=}, {output=}",
            )
            return 1
        common_main(input=input, output=output, template=template)
        return 0
    except Exception as e:
        messagebox.showerror(
            title="An error occured. ",
            message=f"{e}\n"
            f"{traceback.format_exc()}\n\n"
            "Please inform the developer. ",
        )
        raise


def _cli_main():
    sys.exit(cli_main())


def _gui_main():
    sys.exit(gui_main())
