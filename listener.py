#!python3
# -*- coding:utf-8 -*-
import sys
import time
import wx
app=wx.App()
dc=wx.ScreenDC()
screenWidth, screenHeight = wx.GetDisplaySize()

import uiautomation as auto
import os
import os.path
from os import path

import win32gui

import csv

import subprocess


# Speech-To-Text Engine
import speech_recognition as sr
r = sr.Recognizer()

# Text-To-Speech Engine
import pyttsx3
engine = pyttsx3.init()

# Natural Language Processor
import spacy
# Load the installed model "en_core_web_sm"
# From > python -m spacy download en_core_web_sm
# > python -m spacy validate
nlp = spacy.load("en_core_web_sm")


manual_mode = True
tts_enabled = True

yesConfirmations = ['yes', 'yep', 'ok', 'correct', 'right', 'exactly', 'affirmative', "that's it", 'you got it','definitely','sure','of course','not bad', 'yeah','sounds good','got it','nailed it','you nailed it','not bad','good','good to go', 'yes that one', 'this one i like']
INSTRUCTIONS_FOLDER = './instructions/'

# ACCEPT COMMANDS THRU SOCKET --- START
import socket
import threading

connections = []
total_connections = 0

host = '192.168.1.4' # you may change host to your own IP address
port = 6000 # you may change this port to your custom port as long as server and client are a match

# create a socket service
sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
sock.bind((host, port))
sock.listen(5)


class Client(threading.Thread):
    def __init__(self, socket, address, id, name, signal):
        threading.Thread.__init__(self)
        self.socket = socket
        self.address = address
        self.id = id
        self.name = name
        self.signal = signal

    def __str__(self):
        return str(self.id) + " " + str(self.address)

    def run(self):
        # global connections
        # global total_connections
        global commandReceived
        global address
        global sock

        while self.signal:
            try:
                data = self.socket.recv(4096)
            except:
                print("Client " + str(self.address) + " has been disconnected.")
                self.signal = False
                connections.remove(self)
                break
            if data != "":
                print("[Client " + str(self.id) + "] " + str(data.decode("utf-8")))
                commandReceived = str(data.decode("utf-8"))

                # print('RECEIVED ==================> DATA: ', data)
                # print('connections', connections)
                # for client in connections:
                #     if client.id != self.id:
                #         print('RECEIVED ==================', data)
                #         client.socket.sendall(data)





# wait for new connections
def newConnections(socket):
    global sock
    global address
    global connections
    global total_connections

    while True:
        try:
            sock, address = socket.accept()
            connections.append(Client(sock, address, total_connections, "Name", True))
            connections[len(connections) - 1].start()
            # print("Client " + str(connections[len(connections) - 1]) + " is connected.")
            total_connections += 1
        except:
            pass



# ACCEPT COMMANDS THRU SOCKET --- END


def doStep(taskName, instanceCount = 0, run_mode = False):
    time.sleep(1)
    global app
    global dc
    # global run_mode

    # Load Action-Object Data
    ACTION_OBJECT_DB_FILE = './action_object_db/action_object_db.csv'
    with open(ACTION_OBJECT_DB_FILE, newline='') as csvfile:
        action_object_data = list(csv.reader(csvfile, delimiter=','))

    # print(action_object_data)
    # while True:
    # command = getCommand('What is the name of your task?')
    command = taskName.lower()
    doc = nlp(command)

    token_text = [token.text for token in doc]
    # ['This', 'is', 'a', 'text']
    print(token_text)

    token_pos = [token.pos_ for token in doc]
    # ['VERB', 'DET', 'NOUN', 'PUNCT']
    print(token_pos)

    token_dep = [token.dep_ for token in doc]
    # ['ROOT', 'det', 'dobj', 'punct']
    print(token_dep)


    # if commandReceived == "switch to voice":
    #     prompt('Switching to voice mode')
    #     manual_mode = False

    # if commandReceived == "switch to manual mode":
    #     prompt('Switching to manual mode')
    #     manual_mode = True

    if command == "switch to instruction mode":
        prompt('Entering instruction mode...')
        run_mode = False
        instructionMode()
        return '','', 0

    elif command == "help":
        prompt("\n\nYou can say:\nswitch to instruction mode\nwhat is this\nhelp\nexit, shut up, shutdown, goodbye, stop it\nshow instruction list\n\n\nopen a/the/my***\nclick on/the/a\nclick on it\ngo to ***\n\n")
        return '','', 0

    # if commandReceived == "listen":
    #     commandReceived = getCommand("What do you want me to do?")
    #     doStep(commandReceived, 0, False)

    elif command == "show instruction list":
        instructionList = os.listdir(INSTRUCTIONS_FOLDER)
        print('\n\nInstruction List')
        print('======================')
        for instructionIndex, taskName in enumerate(instructionList):
            santizedTaskName = taskName.replace('_',' ').replace('.csv','')
            print(santizedTaskName)
        print('\n\n')
        return '','', 0

    # if commandReceived == "do this":
    #     commandReceived = getCommand("What is the name of the task?")
    #     prompt('Finding instructions from memory...')
    #     runInstructions(commandReceived)

    elif command == "what is this":
        currentMouseX, currentMouseY = auto.GetCursorPos()
        control = auto.ControlFromPoint(currentMouseX, currentMouseY)
        for c, d in auto.WalkControl(control, True, 1000):
            print(c.ControlTypeName, c.ClassName, c.AutomationId, c.BoundingRectangle, c.Name, '0x{0:X}({0})'.format(c.NativeWindowHandle))

        return '','', 0


    elif command == "exit" or command == "shut up" or command == "shutdown" or command == "goodbye" or command == "stop it":
        prompt('Thank you for using virtual employee. Goodbye.')
        sys.exit()




    if len(token_text) > 0:
        verbid = -1
        dobjid = -1

        first_word = token_text[0].lower().strip()
        # override behavior during TYPE command:
        # if first_word == 'type'



        for textid, text in enumerate(token_text):
            if (token_pos[textid] == 'VERB' and (token_dep[textid] == 'ROOT' or token_dep[textid] == 'amod')) or (token_pos[textid] == 'ADJ' and token_dep[textid] == 'amod'):
                verbid = textid
            if token_dep[textid] == 'dobj' or token_dep[textid] == 'pobj' or (token_pos[textid] == 'NOUN' and token_dep[textid] == 'ROOT') or (token_pos[textid] == 'NOUN' and token_dep[textid] == 'advcl') or (token_pos[textid] == 'VERB' and token_dep[textid] == 'advcl') or (token_pos[textid] == 'VERB' and token_dep[textid] == 'advmod') or (token_pos[textid] == 'VERB' and token_dep[textid] == 'xcomp') or (token_pos[textid] == 'VERB' and token_dep[textid] == 'pcomp'):
                dobjid = textid

        if verbid == -1:
            prompt('There is no action in your instruction step. Please start with a verb.')
        else:
            if dobjid == -1:
                prompt('There is no object to act on. Please state a recipient of the action.')
            else:
                action_from_command = str(token_text[verbid].lower().strip())
                object_from_command = str(token_text[dobjid].lower().strip())
                # prompt('Instruction valid.')

                objectInstanceCount = 1  # Counts how many times it should match before selecting the right control

                # Look for built-in actions first
                for action, dobject, subprocess_command, requirement in action_object_data:
                    # print('token_text[verbid]',token_text[verbid])
                    # print('action',action)
                    # print('token_text[dobjid]', token_text[dobjid])
                    # print('dobject', dobject)
                    # print('token_text[verbid] == action', token_text[verbid] == action)
                    # print('token_text[dobjid] == dobject', token_text[dobjid] == dobject)
                    action_from_db = str(action.lower().strip())
                    object_from_db = str(dobject.lower().strip())
                    if action_from_command == action_from_db and object_from_command == object_from_db:
                        # prompt('Instruction found. [' + subprocess_command + ']')

                        if action_from_command == 'open':
                            # print(subprocess_command)
                            subprocess.Popen(subprocess_command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                            time.sleep(2)
                            control = auto.GetRootControl()

                            # Get Target Window
                            detectControls = ['WindowControl', 'PaneControl']
                            for c, d in auto.WalkControl(control, True, 1000):
                                if c.ControlTypeName in detectControls:
                                    # print(c.ControlTypeName, c.ClassName, c.AutomationId, c.BoundingRectangle, c.Name, '0x{0:X}({0})'.format(c.NativeWindowHandle))
                                    if object_from_db.lower() in c.Name.lower():
                                        dc.Clear()
                                        dc.SetPen(wx.Pen("blue"))
                                        dc.SetBrush(wx.Brush("blue", wx.TRANSPARENT))
                                        ControlBoundingRect = c.BoundingRectangle
                                        # TODO: calculate dynamic width/height based on screensize using wx at header
                                        xScale = 1 #1.25
                                        yScale = 1 #1.25
                                        width = (ControlBoundingRect.right * xScale) - (ControlBoundingRect.left * xScale)
                                        height = (ControlBoundingRect.bottom * yScale) - (ControlBoundingRect.top * yScale)
                                        dc.DrawRectangle((ControlBoundingRect.left * xScale),(ControlBoundingRect.top * yScale),width,height)

                                        break

                            control = getActiveWindowControls()

                # Then process all other commands here
                if action_from_command == 'go':
                    prompt(action_from_command + ' to ' + object_from_command)
                    control = getActiveWindowControls()
                    # Get controls of the target window
                    # detectControls = ['TextControl', 'ButtonControl', 'EditControl']
                    possibleControlMatch = []
                    if object_from_command == 'dropdown':
                        possibleControlMatch = ['▾']

                    detectControls = ['HyperlinkControl', 'EditControl', 'ButtonControl', 'ComboBoxControl', 'ListItemControl', 'ListControl', 'TextControl']
                    for c, d in auto.WalkControl(control, True, 1000):
                        # print(c.ControlTypeName, c.ClassName, c.AutomationId, c.BoundingRectangle, c.Name, '0x{0:X}({0})'.format(c.NativeWindowHandle))
                        if object_from_command in c.Name.lower().strip() or object_from_command in possibleControlMatch:
                            if c.ControlTypeName in detectControls:
                            # if object_from_db.lower() in c.Name.lower():
                                dc.Clear()
                                dc.SetPen(wx.Pen("green"))
                                dc.SetBrush(wx.Brush("green", wx.TRANSPARENT))
                                ControlBoundingRect = c.BoundingRectangle
                                xScale = 1 #1.25
                                yScale = 1 #1.25
                                width = (ControlBoundingRect.right * xScale) - (ControlBoundingRect.left * xScale)
                                height = (ControlBoundingRect.bottom * yScale) - (ControlBoundingRect.top * yScale)
                                dc.DrawRectangle((ControlBoundingRect.left * xScale),(ControlBoundingRect.top * yScale),width,height)


                                controlCenterX = int(ControlBoundingRect.left + (width/2))
                                controlCenterY = int(ControlBoundingRect.top + (height/2))
                                auto.SetCursorPos(controlCenterX, controlCenterY)

                                if run_mode == False:
                                    # confirmAction = getCommand('This one?')
                                    commandReceived = ''
                                    commandReceived = getCommand("This one?")
                                    while commandReceived == '':
                                        # print('waiting...')
                                        continue
                                    print('commandReceived',commandReceived)


                                    if commandReceived in yesConfirmations:
                                        break
                                    else:
                                        objectInstanceCount += 1
                                        continue
                                else:
                                    if int(objectInstanceCount) == int(instanceCount):
                                        break
                                    else:
                                        objectInstanceCount += 1
                                        continue

                elif action_from_command == 'click':
                    # prompt(action_from_command + ' on ' + object_from_command)

                    if object_from_command == 'it':
                        currentMouseX, currentMouseY = auto.GetCursorPos()
                        auto.Click(currentMouseX,currentMouseY)
                    else:
                        control = getActiveWindowControls()
                        # Get controls of the target window
                        # detectControls = ['TextControl', 'ButtonControl', 'EditControl']
                        detectControls = ['HyperlinkControl', 'EditControl', 'ButtonControl', 'ComboBoxControl', 'ListItemControl', 'ListControl', 'TextControl']


                        for c, d in auto.WalkControl(control, True, 1000):
                            if object_from_command in c.Name.lower().strip():
                                # print(c.ControlTypeName, c.ClassName, c.AutomationId, c.BoundingRectangle, c.Name, '0x{0:X}({0})'.format(c.NativeWindowHandle))
                                # if c.ControlTypeName in detectControls:
                                # if object_from_db.lower() in c.Name.lower():
                                dc.Clear()
                                dc.SetPen(wx.Pen("green"))
                                dc.SetBrush(wx.Brush("green", wx.TRANSPARENT))
                                ControlBoundingRect = c.BoundingRectangle
                                xScale = 1 #1.25
                                yScale = 1 #1.25
                                width = (ControlBoundingRect.right * xScale) - (ControlBoundingRect.left * xScale)
                                height = (ControlBoundingRect.bottom * yScale) - (ControlBoundingRect.top * yScale)
                                dc.DrawRectangle((ControlBoundingRect.left * xScale),(ControlBoundingRect.top * yScale),width,height)


                                controlCenterX = int(ControlBoundingRect.left + (width/2))
                                controlCenterY = int(ControlBoundingRect.top + (height/2))
                                auto.SetCursorPos(controlCenterX, controlCenterY)

                                if run_mode == False:
                                    commandReceived = getCommand('This one?')
                                    if commandReceived in yesConfirmations:
                                        auto.Click(controlCenterX, controlCenterY)
                                        break
                                    else:
                                        objectInstanceCount += 1
                                        continue
                                else:
                                    print('objectInstanceCount', objectInstanceCount)
                                    print('instanceCount', instanceCount)
                                    if int(objectInstanceCount) == int(instanceCount):
                                        print('It goes here 3')
                                        # auto.Click(controlCenterX, controlCenterY)
                                        currentMouseX, currentMouseY = auto.GetCursorPos()
                                        auto.Click(currentMouseX,currentMouseY)
                                        # time.sleep(2)
                                        break
                                    else:
                                        print('It goes here 2')
                                        objectInstanceCount += 1
                                        continue

                return action_from_command, object_from_command, objectInstanceCount
    return '','', 0

def runInstructions(commandReceived):
    # prompt('Getting all instructions from folder..')
    instructionList = os.listdir(INSTRUCTIONS_FOLDER)
    # print(instructionList)

    matchedInstructionIndex = None

    instructionsFound = False
    for instructionIndex, taskName in enumerate(instructionList):
        santizedTaskName = taskName.replace('_',' ').replace('.csv','')
        if santizedTaskName.lower() == commandReceived.lower():
            instructionsFound = True
            matchedInstructionIndex = instructionIndex
            prompt('Instructions found!')
            break
        # print(santizedTaskName)


    if instructionsFound == False:
        prompt('Sorry, no instructions found in my memory with that name.')
    else:
        # prompt('Running instructions from ' + instructionList[matchedInstructionIndex])
        # Load Instruciton Steps Data
        with open(INSTRUCTIONS_FOLDER + instructionList[matchedInstructionIndex], newline='') as csvfile:
            instruction_steps_data = list(csv.reader(csvfile, delimiter=','))

        for stepno, command, instanceCount, _ in instruction_steps_data:
            # print(stepno, command, instanceCount)
            doStep(command, instanceCount, True)

        prompt('Instruction step execution complete.')

def getActiveWindowControls():

    #Reset control list to just get the currently selected window
    control = auto.GetFocusedControl()
    controlList = []
    while control:
        controlList.insert(0, control)
        control = control.GetParentControl()

    if len(controlList) == 1:
        control = controlList[0]
    else:
        control = controlList[1]
    # Get controls of the target window
    # Draw bounding box of active controls
    # detectControls = ['TextControl', 'ButtonControl', 'EditControl']
    # for c, d in auto.WalkControl(control, True, 1000):
    #     if c.ControlTypeName in detectControls:
    #         print(c.ControlTypeName, c.ClassName, c.AutomationId, c.BoundingRectangle, c.Name, '0x{0:X}({0})'.format(c.NativeWindowHandle))
    #         # if object_from_db.lower() in c.Name.lower():
    #         dc.Clear()
    #         dc.SetPen(wx.Pen("blue"))
    #         dc.SetBrush(wx.Brush("blue", wx.TRANSPARENT))
    #         ControlBoundingRect = c.BoundingRectangle
    #         xScale = 1 #1.25
    #         yScale = 1 #1.25
    #         width = (ControlBoundingRect.right * xScale) - (ControlBoundingRect.left * xScale)
    #         height = (ControlBoundingRect.bottom * yScale) - (ControlBoundingRect.top * yScale)
    #         dc.DrawRectangle((ControlBoundingRect.left * xScale),(ControlBoundingRect.top * yScale),width,height)
    #

    return control

def instructionMode():


    # TASKS
    verifyTaskName = 'no'
    while True:
        taskName = getCommand('What is the name of your task?')
        if isinstance(taskName, str):
            verifyTaskName = getCommand('Task Name: ' + taskName + '. Is this correct?' )
            if verifyTaskName in yesConfirmations:
                # save to file and continue
                prompt("Saving task...")
                instructionMode = True
                instructionFile = taskName
                instructionFile = instructionFile.replace(" ", "_")
                instructionFile = INSTRUCTIONS_FOLDER + instructionFile + '.csv'
                if path.exists(instructionFile)==True:
                    prompt('Saving task cancelled. Task already exists. Please use a different task instead.')
                else:
                    open(instructionFile, 'a').close()
                    break
            elif verifyTaskName == 'cancel':
                prompt("Creating instructions cancelled. Exiting instruction mode.")
                instructionMode = False
                break
            else:
                prompt("Let me ask you again..")


    # STEPS
    instructionCount = 0
    while instructionMode == True:
        instructionCount += 1

        while True:
            # stepName = getCommand('What is step #' + str(instructionCount))
            commandReceived = getCommand('What is step #' + str(instructionCount))
            # print('stepName [', stepName,']')
            if commandReceived == "that is it" or commandReceived == "no more":
                instructionMode = False
                prompt('You have successfully created: ' + taskName + ' instructions.')
                break
            else:
                if isinstance(commandReceived, str):
                    # Do the task then verify
                    action_from_command, object_from_command, objectInstanceCount = doStep(commandReceived, 0, False)
                    if action_from_command == '':
                        prompt("Let me ask you again..")
                    else:
                        commandReceived = getCommand('Step Name: ' + commandReceived + '. Is this correct?' )
                        if commandReceived in yesConfirmations:
                            # save to file and continue
                            stepRow = [instructionCount, commandReceived, objectInstanceCount, 'reserved_parameter']
                            with open(instructionFile, 'a', newline='') as csvFile:
                                writer = csv.writer(csvFile)
                                writer.writerow(stepRow)
                            csvFile.close()

                            prompt("Saving instruction step...")
                            break
                        elif commandReceived == 'cancel':
                            prompt("Creating instructions cancelled. Exiting instruction mode.")
                            instructionMode = False
                            break
                        else:
                            prompt("Let me ask you again..")

def prompt(message):
    print(message)
    if tts_enabled == True:
        engine.say(message)
        engine.runAndWait()

def getCommand(message):
    global commandReceived
    global sock

    prompt(message)
    if manual_mode == True:
        # commandReceived = input('(Manual Mode) Enter a command: ')
        commandReceived = ''
        # create new thread to wait for connections
        newConnectionsThread = threading.Thread(target = newConnections, args = (sock,))
        newConnectionsThread.start()
        # print('====> SERVER LISTENING...')
        # prompt("Command received:" + commandReceived)
        while commandReceived == '':
            # print('waiting...')
            continue
    else:
        with sr.Microphone() as source:
            # prompt("(Voice Mode) Say a command...")
            r.adjust_for_ambient_noise(source)
            commandReceived = r.listen(source)
        try:
            commandReceived = r.recognize_google(commandReceived)
            # Enable command echoing for debugging purposes
            if isinstance(commandReceived, str):
                prompt("Command received:" + commandReceived)
        except:
            return ''

    return commandReceived.lower()

def listen():
    global manual_mode
    global tts_enabled
    global engine
    global r
    global commandReceived


    while True:
        commandReceived = ''
        commandReceived = getCommand("I'm listening...")

        print('commandReceived',commandReceived)

        # MANUAL/BUILT-IN COMMANDS
        if commandReceived != '':
            # if commandReceived == "listen":
            #     commandReceived = getCommand("What do you want me to do?")
            doStep(commandReceived, 0, False)
        else:
            print('no command received.')



if __name__ == '__main__':
    listen()
    # doStep('open browser', 0, False)
    # doStep('go to home', 0, False)
