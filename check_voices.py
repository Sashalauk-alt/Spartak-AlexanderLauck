import win32com.client
tts = win32com.client.Dispatch('SAPI.SpVoice')
voices = tts.GetVoices()
for i in range(voices.Count):
    print(f'{i}: {voices.Item(i).GetDescription()}')