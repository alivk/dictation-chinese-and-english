# Created by Alick for Rocky and Jeremy
import os
import platform
import time

def text_to_speech_windows(text, lang='en'):
    """Use Windows SAPI5 to convert text to speech."""
    import win32com.client
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

def text_to_speech_mac(text, lang='en'):
    """Use macOS 'say' command to convert text to speech."""
    voice_flags = {
        'en': 'Eddy',
        'zh': 'Meijia'  # This is for Mandarin. Replace with 'Sin-ji' for Cantonese.
    }
    os.system(f'say -v {voice_flags.get(lang, "Sandy")} {text}')

def text_to_speech(text, lang='en', delay=None):
    """Converts given text to speech and waits for the given delay."""
    platform_name = platform.system()
    if platform_name == "Windows":
        text_to_speech_windows(text, lang)
    elif platform_name == "Darwin":
        text_to_speech_mac(text, lang)
    else:
        print("Unsupported OS")
        return

    if delay:
        time.sleep(delay)

if __name__ == "__main__":
    # Get buffer timing from user
    try:
        buffer_time = float(input("What are the buffer timing between each dictation words in seconds? "))
    except ValueError:
        print("Invalid input! Please enter a valid number for buffer timing.")
        exit(1)

    # Detect language based on input. Assuming mostly English = 'en' and mostly Chinese = 'zh'.
    lang_input = input("Enter the language (options: 'en' for English, 'zh' for Chinese): ").strip().lower()
    lang = 'en' if lang_input != 'zh' else 'zh'

    # Read sentences from file
    with open('dictation.md', 'r', encoding='utf-8') as file:
        texts = [line.strip() for line in file if line.strip()]

    # Start dictation
    if lang == 'en':
        text_to_speech("Dictation starts now, and every word will give you a short buffer, please write down each word properly before hand over to daddy or mummy to check", lang)
    elif lang == 'zh':
        text_to_speech("听写现在开始，每个单词后都会有短暂的间隔，请正确地写下每个单词，并在交给爸爸或妈妈检查之前核对。", lang)
    
    # Dictate the sentences
    for i, text in enumerate(texts):
        index, sentence = text.split(".", 1)  # Splitting based on the first period to separate the index and sentence
        text_to_speech(index + ".", lang, delay=1)  # Dictate the index with a 1-second buffer
        text_to_speech(sentence.strip(), lang, delay=1)  # Dictate the sentence with a 1-second buffer
        if i != len(texts) - 1:  # Skip "Next one" for the last sentence
            if lang == 'en':
                text_to_speech("Next one", lang, delay=buffer_time)
            elif lang == 'zh':
                text_to_speech("下一个", lang, delay=buffer_time)

    # Start dictation
    if lang == 'en':
        text_to_speech("Please enter index numbers that you would like to listen again, if you want to skip, enter 0", lang)
    elif lang == 'zh':
        text_to_speech("请再次输入您想听的索引号码, 并用逗号分隔,如果您想跳过，请输入0", lang)
   
    
    # Prompt for index numbers to be repeated
    repeat_indexes = input("Please enter index numbers that you would like to listen again, using 0 to skip (e.g. '5,6,7'): ")
    repeat_indexes = [int(i.strip()) for i in repeat_indexes.split(",") if i.strip().isdigit()]
    
    for idx in repeat_indexes:
        if 1 <= idx <= len(texts):
            text = texts[idx - 1]  # Adjust for 0-based indexing
            index, sentence = text.split(".", 1)
            text_to_speech(index + ".", lang, delay=1)
            text_to_speech(sentence.strip(), lang, delay=1)

    # End dictation
    if lang == 'en':
        text_to_speech("This is the end of the dictation, please check again before submit to daddy or mummy", lang)
    elif lang == 'zh':
        text_to_speech("听写结束了，请在提交给爸爸或妈妈之前再次检查。", lang)
   
