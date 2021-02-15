#!/usr/bin/env python3
import os
import zipfile
import shutil
import argparse
import subprocess

audio_extensions = ('.aiff', '.au', '.mid', '.midi', '.mp3', '.m4a', '.mp4', '.wav', '.wma')


def playback_change(speed: float, init_dir: str, quiet: bool, overwrite: bool):
    os.chdir(init_dir)
    init_files = os.listdir()
    number_ppt = 0
    parts = 0
    invalid_ppts = []

    for ppt in init_files:
        if not ppt.endswith('.zip') and not ppt.endswith('.pptx') or ppt == 'base_library.zip':
            continue
        number_ppt += 1

        # Creates folder for extraction
        base = os.path.splitext(ppt)[0]
        if os.path.isdir(base):
            shutil.rmtree(base, ignore_errors=True)
        os.mkdir(base)

        # Extracts the ppt into the folder we created
        new_zip = base + '.zip'
        os.rename(ppt, new_zip)
        with zipfile.ZipFile(new_zip) as myzip:
            for file in myzip.namelist():
                myzip.extract(file, base)
        myzip.close()

        # Checks if media folder exists in ppt
        try:
            os.chdir(os.path.join(init_dir, base, "ppt", "media"))
        except FileNotFoundError:
            invalid_ppts += [f"{ppt}: Media folder missing"]
            os.rename(new_zip, base + '.pptx')
            shutil.rmtree(base, ignore_errors=True)
            continue

        audio_files = [f for f in os.listdir() if f.endswith(audio_extensions)]
        # Catches case where media file exists with no audio files
        if not audio_files:
            invalid_ppts += [f"{ppt}: No audio files"]
            os.chdir(init_dir)
            os.rename(new_zip, base + '.pptx')
            shutil.rmtree(base, ignore_errors=True)
            continue
        # Assumes that numbers start on the 6th character
        sort_audios = sorted(audio_files, key=lambda x: int(os.path.splitext(x)[0][5:]))

        # Loops around media directory, overwriting any audio files based on speed param
        for count, afile in enumerate(sort_audios, start=1):
            print(f"Overwriting slide {count} audio")
            ff_opts = ["-hide_banner", "-stats"]
            if quiet:
                ff_opts = ["-v", "quiet"]
            ff_inp = ["-i", afile]
            ff_out = ["-filter:a", f"atempo={speed}", f"tmp_{afile}"]
            subprocess.run(["ffmpeg"] + ff_opts + ff_inp + ff_out)
            os.remove(afile)
            os.rename(f"tmp_{afile}", afile)
            os.system('cls' if os.name == 'nt' else 'clear')

        # Cleans up the base directory
        os.chdir(init_dir)
        if overwrite:
            os.remove(new_zip)
            shutil.make_archive(f"{base}", 'zip', base)
        else:
            new_ppt_name = f"{base}_{speed}"
            shutil.make_archive(f"{new_ppt_name}", 'zip', base)
            os.rename(f"{new_ppt_name}.zip", f"{new_ppt_name}.pptx")

        os.rename(new_zip, f"{base}.pptx")
        shutil.rmtree(base, ignore_errors=True)
        parts += 1

    for i in invalid_ppts:
        print(i)
    if number_ppt:
        print(f'Processed {parts}/{number_ppt} powerpoints')


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Change the playback speed of powerpoint narration')
    parser.add_argument('speed', type=float, help='desired playback speed')
    parser.add_argument('--quiet', '-q', help='surpress FFmpeg stdout messages', default=False, action='store_true')
    parser.add_argument('--overwrite', '-o', help='overwrite the powerpoint', default=False, action='store_true')
    parser.add_argument('--dir', '-d', help='path to directory containing powerpoints', nargs='?', default=os.getcwd(), type=str)
    args = parser.parse_args()
    if shutil.which("ffmpeg") is None:
        raise OSError("FFmpeg needs to be in your PATH")
    elif not os.path.isdir(args.dir):
        raise OSError("Specified path is not a directory")
    elif not 0.5 <= args.speed <= 100.0:
        raise RuntimeError("Speed must be in the [0.5, 100.0] range")
    playback_change(args.speed, args.dir, args.quiet, args.overwrite)
