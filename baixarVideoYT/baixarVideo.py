import yt_dlp

url = input("Cole a URL do vídeo do YouTube: ")

options = {
    'format': 'bestvideo+bestaudio/best', 
    'outtmpl': '%(title)s.%(ext)s'  
}

try:
    with yt_dlp.YoutubeDL(options) as ydl:
        ydl.download([url])
    print("Download concluído!")
except Exception as e:
    print("Erro ao baixar o vídeo:", e)



""" SE QUISER SÓ O AUDIO 

    options = {
        'format': 'bestaudio/best',
        'outtmpl': '%(title)s.%(ext)s',
        'postprocessors': [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': '192',
        }]
    }
"""