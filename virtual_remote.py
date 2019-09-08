from flask import Flask, render_template
from flask_socketio import SocketIO, emit
import win32gui
import uiautomation as auto
from win32gui import GetWindowText, GetForegroundWindow
import time
import spacy
nlp = spacy.load('en_core_web_sm')

app = Flask(__name__)
socketio = SocketIO(app)

verbs = [
    'click', 'open', 'type', 'go', 'goto', 'move',
    'right click', 'scroll', 'highlight', 'press',
    'tap', 'run', 'close'
    ]

internal_commands = [
    'help', 'listen', 'switch to instruction mode',
    'switch to voice', 'switch to manual mode', 'help me'
    'bye', 'show instruction list','do this',"what is this",
    'exit','shut up','shutdown','goodbye','stop it',
]

yesConfirmations = ['yes', 'yep', 'ok', 'correct', 'right', 'exactly', 'affirmative', "that's it", 'you got it','definitely','sure','of course','not bad', 'yeah','sounds good','got it','nailed it','you nailed it','good','good to go', 'yes that one', 'this one i like']
noConfirmations = ['no', 'nope', 'not ok', 'incorrect', 'wrong', 'not exactly', 'negative', "that's not it", 'of course not', 'no good', 'not that one', "i don't like"]
instructionConfirmations = ['that is it', 'no more', 'cancel']
# modes
# command - time when system is accepting commands
# confirmation - time when system is accepting confirmation

mode = 'command'


import socket
import threading
import sys

def receive(socket, signal):
    while signal:
        try:
            data = socket.recv(4096)
            print(str(data.decode("utf-8")))
        except:
            print("You have been disconnected from the server")
            signal = False
            break


@app.route('/')
def index():
    # return render_template('index.html',**values)
    return render_template('index.html')

@socketio.on('connect')
def test_connect():
    emit('after connect',  {'data':'Lets dance'})

@socketio.on('value changed')
def value_changed(message):
    # emit('update value', message, broadcast=True)
    command = message['data'].lower()
    print(command)

    if mode == 'command':
        checkCommandCompleteness(command)


def checkCommandCompleteness(command):
    print('Received command: [', command,']')
    doc = nlp(command)

    verb = ''
    verb_pos = ''
    verb_tag = ''
    verb_dep = ''

    receiving_object = ''
    receiving_object_pos = ''
    receiving_object_tag = ''
    receiving_object_dep = ''

    receiving_object_specifier = ''
    receiving_object_specifier_pos = ''
    receiving_object_specifier_tag = ''
    receiving_object_specifier_dep = ''

    for token in doc:
        print(token.text, token.pos_, token.tag_, token.dep_)

        if token.text.lower() == 'swish': # command to cancel an incomplete command
            print('')
            print('')
            print('SWISHED!')
            print('')
            print('')
            emit('command detected', 'command detected')
            break
        else:
            # GET VERB
            if (token.pos_ == "VERB" or token.pos_ == "ADJ") and token.text.lower() in verbs:
                # print('WENT INSIDE GETTING VERB')
                verb = token.text
                verb_pos = token.pos_
                verb_tag = token.tag_
                verb_dep = token.dep_

            # GET SPECIFIC OBJECT IF COMPOUND NOUN IS PRESENT, BUT STILL ONLY GET THIS WHEN VERB IS PRESENT
            if (verb_pos == 'VERB' and token.dep_ == 'compound'):
                # print('WENT INSIDE GETTING SPECIFIC OBJECT')
                receiving_object_specifier = token.text
                receiving_object_specifier_pos = token.pos_
                receiving_object_specifier_tag = token.tag_
                receiving_object_specifier_dep = token.dep_

            # GET OBJECT ONLY IF VERB IS PRESENT
            if (verb_pos == 'VERB' and (((token.pos_ == "NOUN" or token.pos_ == "PRON") and (token.dep_ == "dobj" or token.dep_ == "pobj")) or token.dep_ == "advmod")) or (verb_pos == 'ADJ' and token.dep_ == "ROOT" and token.pos_ == "NOUN"):
                # print('WENT INSIDE GETTING OBJECT')
                receiving_object = token.text
                receiving_object_pos = token.pos_
                receiving_object_tag = token.tag_
                receiving_object_dep = token.dep_

    # print('verb:', verb)
    # print('receiving_object:', receiving_object)
    # print('receiving_object_specifier:', receiving_object_specifier)
    # input('pause')

    # send to LISTENER
    host = '192.168.1.4' # you may change host to your own IP address
    port = 6000 # you may change this port to your custom port as long as server and client are a match

    # try to connect to the server app
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.connect((host, port))
    except:
        print("Listener is not active. Connection failed. Cannot find server. Make sure host and port are the same.")
        emit('command detected', 'command detected')
        sys.exit(0)

    # create new thread to wait for data
    receiveThread = threading.Thread(target = receive, args = (sock, True))
    receiveThread.start()

    if verb != '':
        if receiving_object != '':
            if receiving_object_specifier != '':
                print('')
                print('')
                print('DETECTED A COMPLETE COMMAND WITH SPECIFIER!')
                print('')
                print('')
                emit('command detected', 'command detected')
                sock.sendall(str.encode(command))
            else:
                print('')
                print('')
                print('DETECTED A COMPLETE COMMAND!')
                print('')
                print('')
                emit('command detected', 'command detected')

                # message = verb + ' ' + receiving_object
                # sock.sendall(str.encode(message))
                sock.sendall(str.encode(command))
        else:
            print('')
            print('')
            print('No object recipient detected.')
            print('')
            print('')
    else:
        print('command', command)
        print('internal command:', internal_commands)
        if command.lower() in internal_commands or command.lower() in yesConfirmations or command.lower() in noConfirmations or command.lower() in instructionConfirmations:
            sock.sendall(str.encode(command))
            emit('command detected', 'command detected')
        else:
            print('No action commands detected.')



if __name__ == '__main__':
    socketio.run(app, host='192.168.1.4', port=5000)
