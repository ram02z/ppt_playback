# ppt_playback

A command-line tool that utilises ffmpeg in order to overwrite narrated powerpoints with a custom playback speed.

## Requirements

- Python 3.5+

- FFmpeg binary in PATH. See [here](https://github.com/adaptlearning/adapt_authoring/wiki/Installing-FFmpeg) for more information.

## Usage

```properties
python ppt_playback.py N -d="absolute/path/to/dir/"
```

where N is the speed between [0.5, 100.0].

*-d is an optional argument. see below.*

The following message can be shown using the ``h`` or ``--help`` flag.

```properties
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