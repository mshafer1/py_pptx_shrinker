# pptx_shrinker

> Hey, I can't email this powerpoint!

I was told this. And turns out the issue had been that the images used were raw 4k images. 
Causing the pptx file to to be 20 MiB despite having only two images in the deck. That was where the idea emerged.

More recently, I encountered a 30 MiB pptx file with 10 images. Only two of the images were going to be shown full screen.
The rest were already shrunk to be displayed in only a portion of the slide. 
Since the slides are 1920/1080, it does not benefit anyone to keep the larger files for display.

## Requirements

[`ffmepg`](https://ffmpeg.org/ffmpeg.html) must be installed and accessible via PATH

## What does it do?

`shrink-pptx input_file.pptx where_to_save_result.pptx`

Takes in `input_file.pptx`, opens it up (using a system temp directory), and examines each image file.
- If any are over 1920px wide, it downscales them to that resolution (keeping aspect ratio).
- If any file is now over 2 MiB, it is adjusted to have lower quality of image (same resolution)
- If that file is still over 2 MiB, it is discarded and the program iteratively downscales until the image is under 2 MiB.
- If any file is now larger then the original, the original is used.

The resulting image files are placed back the expected locations and the resulting slide deck is saved to `where_to_save_result.pptx`
