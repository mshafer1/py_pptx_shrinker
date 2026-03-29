"""A tool to shrink PowerPoint files by optimizing embedded media."""

import concurrent.futures
import logging
import os.path
import pathlib
import shutil
import subprocess
import tempfile
import zipfile

_logger = logging.getLogger(__name__)
_logger.addHandler(logging.NullHandler())
_MAX_IMAGE_SIZE = 2 * 1024 * 1024  # 2MiB


def _decrements():
    """Generate a sequence of decreasing scale factors for image resizing.

    >>> t = _decrements(); list([next(t) for _ in range(15)])
    [1.0, 0.9, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3, 0.2, 0.1, 0.09, 0.08, 0.07, 0.06, 0.05]

    >>> t = _decrements(); _ = [next(t) for _ in range(15)]; list([next(t) for _ in range(10)])
    [0.04, 0.03, 0.02, 0.01, 0.009, 0.008, 0.007, 0.006, 0.005, 0.004]
    """
    step_size = 1
    current_value = 1.0
    yield current_value
    while True:  # iterate forever
        step = 10 ** (-step_size)
        if current_value - step <= 0:
            step_size += 1
            step = 10 ** (-step_size)
        current_value = round(current_value - step, step_size)
        yield round(current_value, 3)


def _subprocess_run(command: list[str]) -> None:
    """Run a subprocess command and log it."""
    _logger.debug("Running command: %s", " ".join(command))
    subprocess.run(command, check=True, capture_output=True)


def _shrink_image(input_path: pathlib.Path, output_path: pathlib.Path) -> None:
    """Shrink an image using ffmpeg."""
    _subprocess_run(
        [
            "ffmpeg",
            "-i",
            str(input_path),
            "-vf",
            "scale='if(gt(iw,1920),1920,-1):-1'",
            "-y",
            str(output_path),
        ]
    )

    # check if result is over 2MB, if so, try to reduce quality
    if output_path.stat().st_size > _MAX_IMAGE_SIZE:
        _logger.warning(
            "Image %s is larger than 2MB after scaling, trying to reduce quality.", input_path.name
        )
        _subprocess_run(
            [
                "ffmpeg",
                "-i",
                str(input_path),
                "-vf",
                "scale='if(gt(iw,1920),1920,-1):-1'",
                "-q:v",
                "5",
                "-y",
                str(output_path),
            ]
        )

    if output_path.stat().st_size > _MAX_IMAGE_SIZE:
        _logger.warning(
            "Image %s is still larger than 2MB after reducing quality, trying to scale down.",
            input_path.name,
        )
        sizes = iter(_decrements())
        while output_path.stat().st_size > _MAX_IMAGE_SIZE:
            scale_factor = next(sizes)
            _logger.info("Trying scale factor %.3f for image %s...", scale_factor, input_path.name)
            _subprocess_run(
                [
                    "ffmpeg",
                    "-i",
                    str(input_path),
                    "-vf",
                    f"scale='-1:iw*{scale_factor:.3f}'",
                    "-y",
                    str(output_path),
                ]
            )

    # if it's not actually smaller, keep the original
    if output_path.stat().st_size >= input_path.stat().st_size:
        _logger.info(
            "Shrunk image %s is not smaller than original, keeping original.", input_path.name
        )
        shutil.copy(input_path, output_path)


def shrink(input_file: str, output_file: str) -> None:
    """Shrink a PowerPoint file by removing unused media and optimizing images."""
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = pathlib.Path(temp_dir)
        _logger.info("Extracting PowerPoint file to %s...", temp_dir)
        with zipfile.ZipFile(input_file, "r") as zip_ref:
            zip_ref.extractall(temp_dir)
            files = zip_ref.namelist()
            media_files = [f for f in files if f.startswith("ppt/media/")]
            with concurrent.futures.ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
                futures = []
                for media_file in media_files:
                    media_path = temp_path / media_file
                    backup_file = temp_path / (media_file + ".original")
                    shutil.copy(media_path, backup_file)
                    futures.append(executor.submit(_shrink_image, backup_file, media_path))
                concurrent.futures.wait(futures)
        _logger.info("Finished processing media files.")
        with zipfile.ZipFile(output_file, "w") as zip_ref:
            for foldername, _, filenames in os.walk(temp_dir):
                for filename in filenames:
                    if filename.endswith(".original"):
                        continue
                    file_path = os.path.join(foldername, filename)
                    zip_ref.write(file_path, os.path.relpath(file_path, temp_dir))
        _logger.info("Created shrunk PowerPoint file at %s.", output_file)
