import subprocess, os, sys
from build import dur, scene, ass_time
W = os.path.dirname(os.path.abspath(__file__))
src = sys.argv[1]; out = os.path.join(W, "spot-seedance.mp4")
D = dur(src); END = 3.5
scene(os.path.join(W, "end.png"), END, os.path.join(W, "spot_end.mp4"), header=False, full=True)
L = [(0.3, 3, "Enviaba hojas de vida…\\Ny nadie respondía"), (3.1, 6, "Hasta que escribí PERFIL"),
     (6.1, 9, "Un perfil que pasa los filtros\\Ny una IA que filtra las vacantes"), (9.1, 12, "Empezaron a llamarme"),
     (12.1, D - 0.2, "Hoy empiezo\\Nmi nuevo trabajo")]
ass = os.path.join(W, "spot.ass")
with open(ass, "w") as f:
    f.write("[Script Info]\nScriptType: v4.00+\nPlayResX: 1080\nPlayResY: 1920\n\n[V4+ Styles]\n"
            "Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding\n"
            "Style: Sub,Inter Display,62,&H00ECF1F3,&H00ECF1F3,&H003B2A0E,&H283B2A0E,1,0,0,0,100,100,0,0,3,24,0,2,80,80,200,1\n"
            "Style: Tag,Inter,30,&H003B2A0E,&H003B2A0E,&H003DA3E8,&H003DA3E8,1,0,0,0,100,100,2,0,3,12,0,9,40,40,215,1\n\n"
            "[Events]\nFormat: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text\n")
    for a, b, t in L:
        f.write(f"Dialogue: 0,{ass_time(a)},{ass_time(b)},Sub,,0,0,0,,{{\\fad(220,160)\\fscx90\\fscy90\\t(0,240,\\fscx100\\fscy100)}}{t}\n")
    f.write(f"Dialogue: 1,{ass_time(0.3)},{ass_time(D)},Tag,,0,0,0,,{{\\fad(400,300)}}HISTORIA ILUSTRATIVA\n")
X = 0.5
fc = (f"[0:v]scale=1080:1920:force_original_aspect_ratio=increase,crop=1080:1920,fps=30,setsar=1[m];"
      f"[m][2:v]overlay=0:40[mh];[mh]ass={ass}[ms];[1:v]fps=30,setsar=1[e];"
      f"[ms][e]xfade=transition=fade:duration={X}:offset={D - X:.3f}[v];"
      f"[0:a]afade=t=out:st={D + END - X - 2.5:.3f}:d=2.5,apad=pad_dur={END}[a]")
subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", src, "-i", os.path.join(W, "spot_end.mp4"), "-i", os.path.join(W, "header.png"),
                "-filter_complex", fc, "-map", "[v]", "-map", "[a]", "-t", f"{D + END - X:.3f}", "-c:v", "libx264", "-crf", "19",
                "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "192k", "-movflags", "+faststart", out], check=True)
print(out, dur(out))
