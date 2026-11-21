"""Videos con clips animados de Higgsfield: transiciones, subtítulos animados, encabezado de marca y cierre."""
import subprocess, sys, os, re
from build import dur, scene, ass_time

W = os.path.dirname(os.path.abspath(__file__))
C = os.path.join(W, "..", "clips")
FPS, XF = 30, 0.45


def silences(audio):
    out = subprocess.run(["ffmpeg", "-i", audio, "-af", "silencedetect=n=-35dB:d=0.3", "-f", "null", "-"],
                         capture_output=True, text=True).stderr
    return [float(x) for x in re.findall(r"silence_end: ([0-9.]+)", out)]


def seg(src, d, out):
    """Clip of exactly d seconds at 1080x1920. Portrait fills the frame; landscape sits over a blurred copy."""
    cd = dur(src)
    w, h = map(int, subprocess.check_output(["ffprobe", "-v", "error", "-select_streams", "v", "-show_entries",
                                             "stream=width,height", "-of", "csv=p=0", src]).decode().split()[0].split(",")[:2])
    slow = max(1.0, d / (cd - 0.1))
    pts = f"setpts={slow:.4f}*PTS,fps={FPS}"
    if h > w:
        fc = f"[0:v]{pts},scale=1080:1920:force_original_aspect_ratio=increase,crop=1080:1920[v]"
    else:
        fc = (f"[0:v]{pts},split[a][b];[a]scale=1080:1920:force_original_aspect_ratio=increase,crop=1080:1920,"
              f"gblur=sigma=38,eq=brightness=-0.30[bg];[b]scale=1080:-2[fg];[bg][fg]overlay=0:(H-h)/2-80[v]")
    subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", src, "-filter_complex", fc, "-map", "[v]", "-an",
                    "-t", f"{d:.3f}", "-c:v", "libx264", "-pix_fmt", "yuv420p", "-preset", "veryfast", "-crf", "18", out], check=True)


def build(name, audio, lines, visuals, tag=None, tail=3.0):
    """lines: [(subtitle, spoken)]; visuals: [(clip, number_of_lines)] — last visual is the end card."""
    total = dur(audio)
    sil = silences(audio)
    chars = sum(len(s) for _, s in lines)
    # line start times: proportional, snapped to the nearest pause
    starts, acc = [0.0], 0.0
    for _, s in lines[:-1]:
        acc += total * len(s) / chars
        near = min(sil, key=lambda x: abs(x - acc)) if sil else acc
        starts.append(near - 0.1 if abs(near - acc) < 1.3 else acc)
    ends = starts[1:] + [total]
    # visual boundaries
    vb, i = [], 0
    for _, n in visuals:
        vb.append(starts[i]); i += n
    vb.append(total + tail)
    parts = []
    for k, (src, _) in enumerate(visuals):
        d = vb[k + 1] - vb[k] + (XF if k < len(visuals) - 1 else 0)
        out = os.path.join(W, f"{name}_v{k:02d}.mp4")
        if src == "end":
            scene(os.path.join(W, "end.png"), d, out, header=False, full=True)
        else:
            seg(os.path.join(C, src), d, out)
        parts.append((out, d))
    # xfade chain
    inputs, fc, last, t = [], "", "0:v", 0.0
    for k, (p, d) in enumerate(parts):
        inputs += ["-i", p]
    for k in range(1, len(parts)):
        t += parts[k - 1][1] - XF
        lab = f"x{k}"
        fc += f"[{last}][{k}:v]xfade=transition={'fade' if k % 2 else 'smoothup'}:duration={XF}:offset={t:.3f}[{lab}];"
        last = lab
    endstart = vb[-2]
    hdr = len(parts)
    inputs += ["-i", os.path.join(W, "header.png")]
    ass = os.path.join(W, f"{name}.ass")
    with open(ass, "w") as f:
        f.write("[Script Info]\nScriptType: v4.00+\nPlayResX: 1080\nPlayResY: 1920\n\n[V4+ Styles]\n"
                "Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding\n"
                "Style: Sub,Inter Display,58,&H00ECF1F3,&H00ECF1F3,&H003B2A0E,&H283B2A0E,1,0,0,0,100,100,0,0,3,24,0,2,80,80,190,1\n"
                "Style: Tag,Inter,30,&H003B2A0E,&H003B2A0E,&H003DA3E8,&H003DA3E8,1,0,0,0,100,100,2,0,3,12,0,9,40,40,215,1\n\n"
                "[Events]\nFormat: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text\n")
        anim = r"{\fad(200,150)\fscx90\fscy90\t(0,240,\fscx100\fscy100)}"
        for (sub, _), a, b in zip(lines, starts, ends):
            if sub and a < endstart - 0.05:
                f.write(f"Dialogue: 0,{ass_time(a + 0.05)},{ass_time(min(b, endstart))},Sub,,0,0,0,,{anim}{sub}\n")
        if tag:
            f.write(f"Dialogue: 1,{ass_time(0.3)},{ass_time(endstart)},Tag,,0,0,0,,{{\\fad(400,300)}}{tag}\n")
    fc += (f"[{last}][{hdr}:v]overlay=0:40:enable='lt(t,{endstart:.2f})'[h];"
           f"[h]ass={ass}[v]")
    silent = os.path.join(W, f"{name}_silent.mp4")
    subprocess.run(["ffmpeg", "-y", "-loglevel", "error", *inputs, "-filter_complex", fc, "-map", "[v]",
                    "-c:v", "libx264", "-pix_fmt", "yuv420p", "-preset", "medium", "-crf", "20", silent], check=True)
    final = os.path.join(W, f"{name}.mp4")
    subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", silent, "-i", audio, "-filter_complex",
                    f"[1:a]afade=t=in:d=0.05,apad=pad_dur={tail}[a]", "-map", "0:v", "-map", "[a]", "-c:v", "copy",
                    "-c:a", "aac", "-b:a", "192k", "-shortest", "-movflags", "+faststart", final], check=True)
    return final, dur(final)


if __name__ == "__main__":
    if sys.argv[1] == "pro":
        L = [
            ("Bienvenido, bienvenida", "Hola... bienvenido, bienvenida. Qué bueno tenerte aquí."),
            ("Si envías hojas de vida y nadie responde…\\Nno es tu culpa", "Mira, si llevas semanas enviando hojas de vida y el teléfono no suena... no es tu culpa."),
            ("Un filtro automático las descarta\\Nantes de que alguien las lea", "Hoy la mayoría se descartan en un filtro automático, antes de que una persona las lea."),
            ("Perfil Laboral Inteligente\\Nun servicio de Julio Canchano", "Por eso creamos Perfil Laboral Inteligente, un servicio de Julio Canchano."),
            ("Hoja de vida, LinkedIn y certificados\\Nlistos para que te encuentren", "Tomamos tu hoja de vida, tu LinkedIn y tus certificados, y los dejamos listos para que te encuentren."),
            ("Portales de empleo + tablero con IA\\Ncon las vacantes que son para ti", "Configuramos tus portales de empleo... y te montamos un tablero con inteligencia artificial que te trae las vacantes que de verdad son para ti."),
            ("Solo necesitas: tu hoja de vida,\\Nel cargo que buscas y tu ciudad", "Para empezar solo necesitas tres cosas: tu hoja de vida, el cargo que buscas y tu ciudad."),
            ("Un perfil que pasa los filtros\\Ny un sistema que trabaja por ti", "¿Y qué te garantizamos? Un perfil sólido, que pasa los filtros, y un sistema que trabaja por ti todos los días."),
            ("Datos protegidos.\\NSin promesas vacías.", "Con tus datos protegidos, y sin promesas vacías."),
            ("Que buscar empleo vuelva a ser\\Nuna oportunidad", "Queremos que buscar empleo deje de ser una carga... y vuelva a ser una oportunidad."),
            ("Escribe PERFIL al WhatsApp\\N300 528 3369", "Escríbenos PERFIL al WhatsApp: tres cero cero, cinco dos ocho, tres tres seis nueve."),
            ("", "Tu próximo empleo empieza con un perfil que trabaja por ti."),
        ]
        V = [("c01.mp4", 1), ("c09.mp4", 1), ("c05.mp4", 1), ("c02.mp4", 1), ("c03.mp4", 1), ("c07.mp4", 1),
             ("c10.mp4", 1), ("c06.mp4", 1), ("c04.mp4", 1), ("c08.mp4", 1), ("c12.mp4", 1), ("end", 1)]
        print(build("video-bienvenida-pro", os.path.join(W, "bienvenida-elevenlabs.mp3"), L, V))
    else:
        L = [
            ("Durante meses envié hojas de vida…\\Ny nadie respondía", "Durante meses envié hojas de vida... y nadie respondía."),
            ("Empecé a pensar que el problema era yo", "Empecé a pensar que el problema era yo."),
            ("Una noche vi un anuncio:\\NPerfil Laboral Inteligente", "Una noche vi un anuncio: Perfil Laboral Inteligente."),
            ("Escribí PERFIL,\\Nsin muchas esperanzas", "Escribí la palabra PERFIL, sin muchas esperanzas."),
            ("En una videollamada me mostraron por qué\\Nmi hoja de vida no pasaba los filtros", "En una videollamada me mostraron por qué mi hoja de vida no pasaba los filtros."),
            ("En pocos días: un perfil nuevo…", "Y en pocos días tenía un perfil nuevo..."),
            ("…y un tablero con las vacantes\\Nque sí eran para mí", "y un tablero que me traía las vacantes que sí eran para mí."),
            ("Dejé de buscar a ciegas", "Dejé de buscar a ciegas."),
            ("Empezaron a llamarme", "Empezaron a llamarme."),
            ("Y hoy… empiezo en mi nuevo trabajo", "Y hoy... empiezo en mi nuevo trabajo."),
            ("¿Te está pasando a ti también?", "Si a ti también te está pasando,"),
            ("Escribe PERFIL al WhatsApp\\N300 528 3369", "escribe PERFIL al WhatsApp: tres cero cero, cinco dos ocho, tres tres seis nueve."),
            ("", "Tu próximo empleo empieza con un perfil que trabaja por ti."),
        ]
        V = [("c09.mp4", 2), ("c10.mp4", 2), ("c11.mp4", 1), ("c03.mp4", 2), ("c06.mp4", 1), ("c08.mp4", 1),
             ("c12.mp4", 1), ("c01.mp4", 2), ("end", 1)]
        print(build("video-historia", os.path.join(W, "historia-narracion.mp3"), L, V, tag="HISTORIA ILUSTRATIVA"))
