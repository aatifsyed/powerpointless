import argparse
import logging
import sys
from typing import IO, Any, BinaryIO, Optional, TextIO

import argcomplete
from logging_actions import log_level_action
from pptx import Presentation as presentation
from pptx.presentation import Presentation
from rich.logging import RichHandler

from .core import create_subtitles, extract_subtitles

logger = logging.getLogger(__name__)


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
    logger.addHandler(RichHandler())

    parser = argparse.ArgumentParser(description="Work with powerpoint subtitles. ")
    parser.add_argument(
        "-l", "--log-level", action=log_level_action(logger), default="debug"
    )

    subparsers = parser.add_subparsers(dest="subcommand", required=True)
    create_subtitles_parser = subparsers.add_parser(
        "create-subtitles", help="Create a powerpoint from a plain text file. "
    )

    create_subtitles_parser.add_argument(
        "-t",
        "--template",
        type=argparse.FileType(mode="rb"),
        help="Look at the first slide master from this file. "
        "Look at the first provided layout in that master. "
        "Create a new slide by populating the first placeholder in that layout. "
        "Defaults to an internal layout. ",
        default=None,
    )
    create_subtitles_parser.add_argument(
        "-i",
        "--input",
        type=MyFileType("r"),
        help="A new slide will be created for each line in this file. ",
        required=True,
    )
    create_subtitles_parser.add_argument(
        "-o",
        "--output",
        type=MyFileType("wb"),
        help="Write resulting presentation to this file. ",
        required=True,
    )

    extract_subtitles_parser = subparsers.add_parser(
        "extract-subtitles", help="Convert a powerpoint into a plain text file. "
    )
    extract_subtitles_parser.add_argument(
        "-i",
        "--input",
        type=MyFileType("rb"),
        help="A new line will be created for each textbox in each slide in this powerpoint. ",
        required=True,
    )
    extract_subtitles_parser.add_argument(
        "-o",
        "--output",
        type=MyFileType("w"),
        help="Text file containing the powerpoint contents. ",
        required=True,
    )

    argcomplete.autocomplete(parser)
    args = parser.parse_args()

    logger.debug(args)

    if args.subcommand == "create-subtitles":
        input: TextIO = args.input
        output: BinaryIO = args.output
        template: Optional[BinaryIO] = args.template

        try:
            tmpl: Presentation = presentation(template)
        except Exception as e:
            raise RuntimeError("Couldn't open file as a powerpoint") from e

        create_subtitles(template=tmpl, lines=input.readlines()).save(output)

        return 0

    elif args.subcommand == "extract-subtitles":
        input: BinaryIO = args.input
        output: TextIO = args.output

        try:
            source: Presentation = presentation(input)
        except Exception as e:
            raise RuntimeError("Coulnd't open file as a powerpoint") from e

        output.write("\n".join(extract_subtitles(source)))

        return 0

    else:
        raise RuntimeError("Unreachable")


def _cli_main():
    sys.exit(cli_main())
