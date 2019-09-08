from flask import Flask, render_template
from flask_socketio import SocketIO, emit

app = Flask(__name__)

socketio = SocketIO(app)

values = {
    'slider1': 25,
    'slider2': 0,
}

@app.route('/')
def index():
    return render_template('index1.html', **values)

@socketio.on('connect')
def test_connect():
    emit('after connect',  {'data':'Lets dance'})

@socketio.on('value changed')
def value_changed(message):
    values[message['who']] = message['data']
    emit('update value', message, broadcast=True)
    print('Message Recieved', message)


if __name__ == '__main__':
    socketio.run(app, host='192.168.22.96', port=5000)
