"""Command-line interface for pptx_shrinker."""

import logging
import pathlib

import click

import pptx_shrinker

_logger = logging.getLogger(__name__)
_logger.addHandler(logging.NullHandler())


@click.command()
@click.argument("input_file", type=click.Path(exists=True, path_type=pathlib.Path, dir_okay=False))
@click.argument("output_file", type=click.Path(dir_okay=False, path_type=pathlib.Path))
def cli(input_file: pathlib.Path, output_file: pathlib.Path) -> None:
    """Shrink a PowerPoint file by scaling images down to 1920px wide and reducing quality.

    \b
    If the image is larger than 1920x1080,
      it will be scaled down to fit within 1920x1080 while maintaining aspect ratio.
    If the resulting image is larger than 2 MiB,
      the quality will be reduced to try to get it under 2 MiB.
    If the resulting image is still larger than 2 MiB,
      it will be scaled down further in increments of 10% until it is under 2 MiB.

    \b
    If the resulting image is not smaller than the original,
      the original will be kept.
    """  # noqa: D301 \b is used to prevent click from reformatting the help text
    logging.basicConfig(level=logging.INFO)
    pptx_shrinker.shrink(input_file, output_file)
    print(f"Shrunk PowerPoint saved to {output_file}")
    print("Original size: {:.2f} MiB".format((input_file.stat().st_size) / (1024 * 1024)))
    print("Shrunk size: {:.2f} MiB".format((output_file.stat().st_size) / (1024 * 1024)))


if __name__ == "__main__":
    cli()
