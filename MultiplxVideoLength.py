import sys
import subprocess
import json
import time
import os
import win32com.client

from pymediainfo import MediaInfo


def get_ffprobe_version():
    return subprocess.check_output(["ffprobe", "-version"]).splitlines()[0].decode()


def ffprobe(vid_file_path):
    """
    Give a json from ffprobe command line
    https://github.com/gbstack/ffprobe-python/blob/master/ffprobe/ffprobe.py

    :param vid_file_path: The absolute (full) path of the video file, string.
    :return: json from ffprobe
    """

    #
    # Command line use of 'ffprobe':
    #
    # ffprobe -loglevel quiet -print_format json \
    #         -show_format    -show_streams \
    #         video-file-name.mp4
    #
    # man ffprobe # for more information about ffprobe
    #

    # maybe this is only working on windows
    command = ["ffprobe", "-loglevel", "quiet", "-print_format", "json", "-show_format", "-show_streams",
               vid_file_path]

    # shell=True?
    with subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE) as proc:
        out, err = proc.communicate()
        return json.loads(out)


def ffprobe_duration(vid_file_path):
    if not ffprobe_available:
        return False

    _json = ffprobe(vid_file_path)
    # Video's duration in seconds, return a float number

    if 'format' in _json:
        if 'duration' in _json['format']:
            return float(_json['format']['duration'])

    if 'streams' in _json:
        # commonly stream 0 is the video
        for s in _json['streams']:
            if 'duration' in s:
                return float(s['duration'])
    return False


def duration(vid_file_path):
    media_info = MediaInfo.parse(vid_file_path)
    if info_print:
        media_info_print(media_info)
    duration_in_ms = media_info.tracks[0].duration
    if duration_in_ms:
        duration_in_s = duration_in_ms/1000
        return duration_in_s
    if os.path.getsize(vid_file_path) < 100:
        global corrupt_files
        corrupt_files.append(vid_file_path)
        raise NameError('I found no duration in file ' + vid_file_path + ' because the file is corrupt')

    print("If ffprobe is available try with it")
    duration_in_s = ffprobe_duration(vid_file_path)
    if duration_in_s:
        return duration_in_s

    # if everything didn't happen,
    # we got here because no single 'return' in the above happen.
    raise NameError('I found no duration in file ' + vid_file_path)


def media_info_print(media_info):
    def track_extract(t, need):
        ttype = t.track_type
        if need in t.to_data():
            data = t.to_data()[need]
            if need is "other_duration":
                data = data[3][0:8]
            if need is "other_format" or need is "other_channel_s":
                data = data[0]
            return ttype, need, " "*(22-len(need)), data
        return ttype, "not found"

    for track in media_info.tracks:
        # to see all data uncomment
        # for k in track.to_data().keys():
        # print("{}.{}={}".format(track.track_type, k, track.to_data()[k]))
        if track.track_type == 'Video':
            print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
            for info in ["width", "height", "duration", "other_duration", "other_format", "codec_id"]:
                print("{} {}{}{}".format(*track_extract(track, info)))
            print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
        elif track.track_type == 'Audio':
            print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
            for info in ["format", "codec_id", "channel_s", "other_channel_s"]:
                print("{} {}{}{}".format(*track_extract(track, info)))
            print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")


def seconds_to_d_h_m_s(temp_seconds, lang):
    i_seconds = divmod(temp_seconds, 60)[1]
    i_minutes1 = divmod(temp_seconds, 60)[0]
    i_hours1 = divmod(i_minutes1, 60)[0]
    i_minutes2 = divmod(i_minutes1, 60)[1]
    i_days = divmod(i_hours1, 24)[0]
    i_hours2 = divmod(i_hours1, 24)[1]
    if int(i_days) > 0:
        diff_time = str(int(i_days)) + " " + lang[0] + ", " + str(int(i_hours2)) + " " + lang[1] + ", " + str(
            int(i_minutes2)) + " " + lang[2] + ", " + str(round(i_seconds)) + " " + lang[3]
    elif int(i_hours2) > 0:
        diff_time = str(int(i_hours2)) + " " + lang[1] + ", " + str(int(i_minutes2)) + " " + lang[2] + ", " + str(
            round(i_seconds)) + " " + lang[3]
    elif int(i_minutes2) > 0:
        diff_time = str(int(i_minutes2)) + " " + lang[2] + ", " + str(round(i_seconds)) + " " + lang[3]
    else:
        diff_time = str(round(i_seconds)) + " " + lang[3]

    return diff_time


def has_video_endings(file):
    media_endings = [".mp4", ".mkv", ".wmv", ".mov"]
    return [file for ending in media_endings if file.lower().endswith(ending)]


def has_link_endings(file):
    link_endings = [".lnk", ".url"]
    return [file for ending in link_endings if file.lower().endswith(ending)]


def resolve_link(file):
    # TODO This is working for Windows .lnk Links but does every link get resolved? Other os?
    if not has_link_endings(file):
        return False
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(file)
    real_file = shortcut.Targetpath
    print("link: " + file + " ------- > " + real_file)
    return file_or_path(real_file)


def progress_file(file):
    global total_counter
    if has_video_endings(file):
        video_file = file
        global total_duration
        global counter
        try:
            file_duration = duration(video_file)
        except NameError as err:
            print(err)
            return False
        print(file_duration)

        total_duration += file_duration
        counter += 1
        total_counter += 1
        return True

    if not resolve_link(file):
        print("This file has no video ending and is not a link")
        total_counter += 1
        return False


def progress_path(path):
    # walk should not go down to symlinks
    for (dirpath, dirnames, filenames) in os.walk(path):
        if len(filenames) > 0:
            get_files[dirpath] = filenames
            for filename in filenames:
                print(filename + " ------- > " + languages[user_language][4] + ": " + dirpath)
                file_or_path(os.path.join(dirpath, filename))
    return True


def file_or_path(x):
    if type(x) != str or x is False:
        return False

    if not os.path.exists(x):
        print(languages[user_language][5] + ": " + x)
        return False

    if os.path.isfile(x):
        progress_file(x)
    elif os.path.isdir(x):
        progress_path(x)
    else:
        print(languages[user_language][6])
        return False
    return True


if __name__ == "__main__":

    get_files = {}
    counter = 0
    total_counter = 0
    total_duration = 0
    corrupt_files = []
    info_print = False
    en = ["days", "hours", "minutes", "seconds", "in directory", "This path does not exist",
          "This is not a file or a path", "Enter the directory", "More directories",
          "Enter the additional directory", "corrupt file", "Quit"]
    de = ["Tage", "Stunden", "Minuten", "Sekunden", "im Verzeichnis", "Dieser Pfad existiert nicht",
          "Das ist keine Datei und kein Pfad", "Gib das Verzeichnis an", "Noch mehr Verzeichnisse",
          "Gib das zusätzliche Verzeichnis an", "Korrupte Datei", "Beenden"]
    languages = {"en": en, "de": de}
    user_language = "en"

    try:
        print(get_ffprobe_version())
        ffprobe_available = True
    except FileNotFoundError:
        print("ffprobe not found - skip")
        ffprobe_available = False

    if len(sys.argv) > 1:
        files_paths = sys.argv[1:]
    else:
        files_paths = []
        raw_input = input(languages[user_language][7] + ": ")
        files_paths.append(raw_input)
        while True:
            raw_input = input(languages[user_language][8] + "? (y/n): ").strip()
            if raw_input == "y":
                files_paths.append(input(languages[user_language][9] + ": "))
            else:
                break
    print(files_paths)
    start_time = time.time()
    for item in files_paths:
        file_or_path(item)
    print("--- Duration Command: %s seconds ---" % (time.time() - start_time))
    print(total_duration)
    print(seconds_to_d_h_m_s(total_duration, languages[user_language]))
    print(str(counter) + " of " + str(total_counter) + " files were videos")
    print("\n".join([""] + [languages[user_language][10] + ": " + a for a in corrupt_files] +
                    [""] if len(corrupt_files) > 0 else []))
    while input(languages[user_language][11] + "?: y ") != "y":
        pass
