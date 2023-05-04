from flask import Flask, render_template, url_for, request, jsonify, redirect, session, g, Response
import os
from flask_mysqldb import MySQL,MySQLdb
from flask_bootstrap import Bootstrap
import bcrypt
from object_detection import *
from camera_settings import *
import gradio as gr
import openai
from gtts import gTTS
import win32com.client
import pythoncom
from pydub import AudioSegment
import threading



currentlocation = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)
Bootstrap(app)
app.static_folder = 'static'

openai.api_key = "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"

messages = [{"role": "system", "content": "You are a helpful assistant."}]


app.secret_key = os.urandom(24)

app.config['MYSQL_HOST'] = 'localhost'
app.config['MYSQL_USER'] = 'root'
app.config['MYSQL_PASSWORD'] = ''
app.config['MYSQL_DB'] = 'flaskdb'
app.config['MYSQL_CURSORCLASS'] = 'DictCursor'
mysql = MySQL(app)

check_settings()
VIDEO = VideoStreaming()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/about')
def about():
    return render_template('about.html')

@app.route('/contact')
def contact():
    return render_template('contact.html')

@app.route('/course1')
def course1():
    return render_template('course1.html')

@app.route('/course2')
def course2():
    return render_template('course2.html')

@app.route('/course2/video_feed')
def video_feed():
    """
    Video streaming route.
    """
    return Response(
        VIDEO.show(),
        mimetype="multipart/x-mixed-replace; boundary=frame"
    )

# * Button requests
@app.route("/course2/request_preview_switch")
def request_preview_switch():
    VIDEO.preview = not VIDEO.preview
    print("*"*10, VIDEO.preview)
    return "nothing"


@app.route("/course2/request_flipH_switch")
def request_flipH_switch():
    VIDEO.flipH = not VIDEO.flipH
    print("*"*10, VIDEO.flipH)
    return "nothing"


@app.route("/course2/request_model_switch")
def request_model_switch():
    VIDEO.detect = not VIDEO.detect
    print("*"*10, VIDEO.detect)
    return "nothing"


@app.route("/course2/request_exposure_down")
def request_exposure_down():
    VIDEO.exposure -= 1
    print("*"*10, VIDEO.exposure)
    return "nothing"


@app.route("/course2/request_exposure_up")
def request_exposure_up():
    VIDEO.exposure += 1
    print("*"*10, VIDEO.exposure)
    return "nothing"


@app.route("/course2/request_contrast_down")
def request_contrast_down():
    VIDEO.contrast -= 4
    print("*"*10, VIDEO.contrast)
    return "nothing"


@app.route("/course2/request_contrast_up")
def request_contrast_up():
    VIDEO.contrast += 4
    print("*"*10, VIDEO.contrast)
    return "nothing"


@app.route("/course2/reset_camera")
def reset_camera():
    STATUS = reset_settings()
    print("*"*10, STATUS)
    return "nothing"

@app.route('/courses')
def courses():
    return render_template('courses.html')

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        session.pop('user', None)

        if request.form['password'] == 'password':
            session['user'] = request.form['username']
            return redirect(url_for('protected'))

    return render_template("login.html")

@app.route('/protected')
def protected():
    if g.user:
        return render_template("protected.html", user=session['user'])
    return redirect(url_for('login'))

@app.before_request
def before_request():
    g.user = None

    if 'user' in session:
        g.user = session['user']

@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for("login"))

@app.route('/register', methods=['GET', 'POST'])
def register():
        if request.method == 'GET':
            return render_template("register.html")
        else:
            name = request.form['name']
            email = request.form['email']
            password = request.form['password'].encode('utf-8')
            hash_password = bcrypt.hashpw(password, bcrypt.gensalt())

            cur = mysql.connection.cursor()
            cur.execute("INSERT INTO users (name, email, password) VALUES (%s,%s,%s)", (name, email, hash_password,))
            mysql.connection.commit()
            session['name'] = request.form['name']
            session['email'] = request.form['email']
        return redirect(url_for('index'))

def decipher(audio):
    global messages

    # Using openAI's speech to text model
    audio_file = open(audio, "rb")
    transcript = openai.Audio.transcribe("whisper-1", audio_file)

    messages.append({"role": "user", "content": transcript["text"]})

    response =  openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=messages
    )

    system_message = response["choices"][0]["message"]["content"]
    pythoncom.CoInitialize()
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(system_message)
    # myobj = gTTS(text=system_message, lang=language, slow=False)
    # myobj.save("welcome.mp3")
    # # Playing the converted file
    # os.system("start welcome.mp3")
    messages.append({"role": "assistant", "content": system_message},)

    chat_transcript = ""
    for message in messages:
        if message['role'] != 'system':
            chat_transcript += message['role'] + ": " + message['content'] + "\n\n"

    return chat_transcript

@app.route('/voice')
def voice():
    if request.method == 'POST':
        audio = request.files['audio'].read()
        response = decipher(audio)
        return response

    # create Gradio interface
    iface = gr.Interface(
        fn=decipher,
        inputs=gr.inputs.Audio(source='microphone', type='filepath'),
        outputs='text',
        capture_session=True,
        title='Interactive Voice Assistant',
        allow_flagging=False,
        flagging_dir=None
    )

    def launch_gradio_interface():
        iface.launch(share=True, debug=False, inline=False)
    gradio_thread = threading.Thread(target=launch_gradio_interface)
    gradio_thread.start()

    return render_template('voice_ass.html', gradio_html="")


if __name__ == '__main__':
    app.run(debug=True)
