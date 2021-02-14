# ppt_playback

A command-line tool that utilises ffmpeg in order to overwrite narrated powerpoints with a custom playback speed.

## Requirements

- Python 3.5+

- FFmpeg binary in PATH. See [here](https://github.com/adaptlearning/adapt_authoring/wiki/Installing-FFmpeg) for more information.

## Usage

```python
python ppt_playback.py -d="absolute/path/to/dir/" N
```

where N a speed value between [0.5, 100.0].

For each powerpoint in the directory, a new powerpoint will be created unless the ``-o`` flag is used.

*-d is an optional argument. defaults to current working directory.*

The following message can be shown using the ``h`` or ``--help`` flag.

```console
usage: ppt_playback.py [-h] [--quiet] [--overwrite] [--dir [DIR]] speed

Change the playback speed of powerpoint narration

positional arguments:
  speed                 desired playback speed

optional arguments:
  -h, --help            show this help message and exit
  --quiet, -q           surpress FFmpeg stdout messages
  --overwrite, -o       overwrite the powerpoint
  --dir [DIR], -d [DIR]
```

### Tested on the following platforms

- Ubuntu 20.04/20.10

- Windows 10

Should also work on any platform that meets the [requirements](#Requirements).

Inspired by the [PowerPointAudio-Extractor](https://github.com/Tortar/PowerPointAudio-Extractor) script.
