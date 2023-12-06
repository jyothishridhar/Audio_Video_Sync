import moviepy.editor as mp
import numpy as np
from scipy.stats import skew
import pandas as pd
import requests
from io import BytesIO
import xlsxwriter

print("Before importing moviepy")
import moviepy.editor as mp
print("After importing moviepy")


def download_file(url, dest_path):
    response = requests.get(url)
    with open(dest_path, 'wb') as file:
        file.write(response.content)

def sync_and_report(video_url, output_path_sync, output_path_report, delay_offset_unsync=0.5):
    # Download the video file
    video_dest_path = "video.mp4"
    download_file(video_url, video_dest_path)

    # Load the video file
    video = mp.VideoFileClip(video_dest_path)

    # Extract the audio and video tracks
    audio = video.audio

    # Get the frame rate and duration of the video
    video_fps = video.fps
    video_duration = video.duration

    # Resample the audio track to match the frame rate of the video track
    audio_resampled = audio.set_fps(video_fps)

    # Get the number of frames for the resampled audio track
    audio_frame_count = int(audio_resampled.duration * audio_resampled.fps)

    # Get the number of frames for the video track
    video_frame_count = int(video_duration * video_fps)

    # Calculate the delay values and corresponding time for each frame for sync
    video_delay_values_sync = []
    video_frame_times_sync = []
    audio_delay_values_sync = []
    audio_frame_times_sync = []

    for frame_number in range(video_frame_count):
        frame_time = frame_number / video_fps
        frame_audio = audio_resampled.subclip(frame_time, frame_time + 1 / video_fps)
        frame_video_delay = 1 / video_fps
        frame_audio_delay = frame_audio.duration
        video_delay_values_sync.append(frame_video_delay)
        video_frame_times_sync.append(frame_time)
        audio_delay_values_sync.append(frame_audio_delay)
        audio_frame_times_sync.append(frame_time)

    # Calculate the skewness of the audio delay values for sync
    audio_skewness_sync = skew(audio_delay_values_sync)

    # Create a DataFrame for synchronization report
    report_data_sync = {
        'Sync_Frame': list(range(1, video_frame_count + 1)),
        'Sync_Duration_video': video_delay_values_sync,
        'Sync_Time': video_frame_times_sync,
        'Sync_Duration_audio': audio_delay_values_sync,
        'Sync_Summary': [sum(abs(video_delay - audio_delay) for video_delay, audio_delay in zip(video_delay_values_sync, audio_delay_values_sync))]
    }

    # Make sure all columns have the same length
    max_length = max(len(report_data_sync[col]) for col in report_data_sync)
    for col in report_data_sync:
        report_data_sync[col] += [np.nan] * (max_length - len(report_data_sync[col]))

    report_df_sync = pd.DataFrame(report_data_sync)

    # Create a Pandas Excel writer for the report
    excel_writer = pd.ExcelWriter(output_path_report, engine='xlsxwriter')

    # Write the DataFrame to the Excel file for synchronization report
    report_df_sync.to_excel(excel_writer, sheet_name='Sync_Frames', index=False)

    # Write additional information to the Excel file for synchronization report
    worksheet_sync = excel_writer.sheets['Sync_Frames']
    worksheet_sync.write(video_frame_count + 3, 0, 'Audio_Video_delay_seconds:')
    worksheet_sync.write(video_frame_count + 3, 1, report_data_sync['Sync_Summary'][0])
    worksheet_sync.write(video_frame_count + 6, 0, 'Audio Skewness:')
    worksheet_sync.write(video_frame_count + 6, 1, audio_skewness_sync)
    worksheet_sync.write(video_frame_count + 8, 0, 'Number of frames in audio:')
    worksheet_sync.write(video_frame_count + 8, 1, audio_frame_count)
    worksheet_sync.write(video_frame_count + 9, 0, 'Number of frames in video:')
    worksheet_sync.write(video_frame_count + 9, 1, video_frame_count)

    # Save the synchronization Excel file
    excel_writer.save()

    # Write the synchronized video to the output file
    synced_video = video.set_audio(audio)
    synced_video.write_videofile(output_path_sync, codec='libx264', audio_codec='aac')

    # Load the synchronized video file
    video_synced = mp.VideoFileClip(output_path_sync)
    audio_synced = video_synced.audio

    # Calculate the delay values and corresponding time for each frame for unsync
    video_delay_values_unsync = []
    video_frame_times_unsync = []
    audio_delay_values_unsync = []
    audio_frame_times_unsync = []

    for frame_number in range(video_frame_count):
        frame_time = frame_number / video_fps
        t_start = frame_time + delay_offset_unsync
        t_end = frame_time + 1 / video_fps + delay_offset_unsync

        if t_start <= video_duration:
            frame_audio = audio_resampled.subclip(max(t_start, 0), min(t_end, video_duration))
            frame_video_delay = 1 / video_fps
            frame_audio_delay = frame_audio.duration
            video_delay_values_unsync.append(frame_video_delay)
            video_frame_times_unsync.append(frame_time)
            audio_delay_values_unsync.append(frame_audio_delay)
            audio_frame_times_unsync.append(frame_time)

    # Calculate the total delay for video and audio separately for unsync
    total_video_delay_unsync = sum(video_delay_values_unsync)
    total_audio_delay_unsync = sum(audio_delay_values_unsync)

    # Calculate the skewness of the audio delay values for unsync
    audio_skewness_unsync = skew(audio_delay_values_unsync)

    # Create a DataFrame with the delay and time for each video frame for unsync
    report_data_unsync = {
        'Unsync_Frame': list(range(1, len(video_delay_values_unsync) + 1)),
        'Unsync_Duration_video': video_delay_values_unsync,
        'Unsync_Time': video_frame_times_unsync,
        'Unsync_Duration_audio': audio_delay_values_unsync,
        'Unsync_Summary': [total_video_delay_unsync]
    }

    # Make sure all columns have the same length
    max_length_unsync = max(len(report_data_unsync[col]) for col in report_data_unsync)
    for col in report_data_unsync:
        report_data_unsync[col] += [np.nan] * (max_length_unsync - len(report_data_unsync[col]))

    report_df_unsync = pd.DataFrame(report_data_unsync)

    # Add unsync data to the existing Excel file
    with pd.ExcelWriter(output_path_report, engine='openpyxl', mode='a') as writer:
        # Write the DataFrame to the Excel file for unsynchronization report
        report_df_unsync.to_excel(writer, sheet_name='Unsync_Frames', index=False)

        # Write additional information to the Excel file for unsynchronization report
        worksheet_unsync = writer.sheets['Unsync_Frames']
        worksheet_unsync.cell(row=video_frame_count + 4, column=1, value='Audio_Video_delay_seconds:')
        worksheet_unsync.cell(row=video_frame_count + 4, column=2, value=total_video_delay_unsync)
        worksheet_unsync.cell(row=video_frame_count + 7, column=1, value='Audio Skewness:')
        worksheet_unsync.cell(row=video_frame_count + 7, column=2, value=audio_skewness_unsync)
        worksheet_unsync.cell(row=video_frame_count + 9, column=1, value='Number of frames in audio:')
        worksheet_unsync.cell(row=video_frame_count + 9, column=2, value=audio_frame_count)
        worksheet_unsync.cell(row=video_frame_count + 10, column=1, value='Number of frames in video:')
        worksheet_unsync.cell(row=video_frame_count + 10, column=2, value=video_frame_count)

    # Clean up the temporary video file
    os.unlink(video_dest_path)

# Git LFS URL for the video
video_url_input = "https://github.com/jyothishridhar/Audio_Video_Sync/raw/master/sync_video.mp4"
output_path_sync_input = 'sync_video.mp4'
output_path_report_input = 'report_combined.xlsx'

# Call the function with input parameters
sync_and_report(video_url_input, output_path_sync_input, output_path_report_input)

