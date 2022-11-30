import socket
import threading
import tkinter
from tkinter import scrolledtext
from tkinter import simpledialog 
import time
import easygui
from tkinter import filedialog as fd
from tkinter import *
# import tqdm
import os

SEPARATOR = "<SEPARATOR>"
BUFFER_SIZE = 4096 # send 4096 bytes each time step

HOST='127.0.0.1'
PORT=9090

class Client:

    def __init__(self,host, port):

        self.s=socket.socket(socket.AF_INET,socket.SOCK_STREAM)
        self.s.connect((host,port))

        msg=tkinter.Tk()
        msg.withdraw()

        self.username=simpledialog.askstring("Username", "Please choose a Username",parent=msg)
        # self.email=simpledialog.askstring("Email", "Enter your email ID here",parent=msg)
        

        self.ui_complete=False

        self.running=True
        uiThread=threading.Thread(target=self.ui)
        RecvThread=threading.Thread(target=self.receive)

        uiThread.start()
        RecvThread.start()
    def ui(self):
        self.window=tkinter.Tk()
        self.window.geometry('650x660')
        self.win=tkinter.PanedWindow(self.window)           #creating two sections in the main window
        self.win.configure(bg="lightgray")
        self.win.pack(fill=BOTH, expand=1)

        self.connected_users = tkinter.PanedWindow(self.win)   ##connected user left window
        self.connected_users.pack(padx=5,pady=5)
        self.win.add(self.connected_users) 

        self.connected_users_label=tkinter.Label(self.connected_users,text="Connected Users",bg="lightgray")
        self.connected_users_label.config(font=("Arial",12))
        self.connected_users_label.pack(padx=20,pady=5)

        self.users=scrolledtext.ScrolledText(self.connected_users)
        self.users.pack(padx=0,pady=0)
        self.users.config(state='disabled',width=20)
        self.chat_area = tkinter.PanedWindow(self.win)
        self.win.add(self.chat_area)

        
        self.chat_area_label=tkinter.Label(self.chat_area,text="Chat:",bg="lightgray")  ##chat area right window pane
        self.chat_area_label.config(font=("Arial",12))
        self.chat_area_label.pack(padx=20,pady=5)

        self.msgs=scrolledtext.ScrolledText(self.chat_area)
        self.msgs.pack(padx=20,pady=5)
        self.msgs.config(state='disabled')

        self.msgs_label=tkinter.Label(self.chat_area,text="Message:",bg="lightgray")
        self.msgs_label.config(font=("Arial",12))     #giving a label to the message box area
        self.msgs_label.pack(padx=20,pady=5)

        self.msg_input=tkinter.Text(self.chat_area,height=3)
        self.msg_input.pack(padx=20,pady=5)

        self.send_btn=tkinter.Button(self.chat_area,text="Send",command=self.write_msg)
        self.send_btn.config(font=("Arial",12))       ##button to send the file
        self.send_btn.pack(padx=20,pady=5)

        self.file_btn=tkinter.Button(self.chat_area,text="Attach",command=self.send_file)
        self.file_btn.config(font=("Arial",12))       ##button to send the message
        self.file_btn.pack(padx=20,pady=5)

        self.exit_btn=tkinter.Button(self.chat_area,text="EXIT CHAT",command=self.exit)
        self.exit_btn.config(font=("Arial",12))            ##exit button to leave chat
        self.exit_btn.pack(padx=20,pady=5)

        self.ui_complete=True
        self.window.protocol("WM_DELETE_WINDOW",self.exit)
        self.win.mainloop()
    
    def send_file(self):
        print("pressed the send file button")
        fileName=easygui.fileopenbox()
        fileSize = os.path.getsize(fileName)
        print('got filename and size')
        self.s.send("!FTP".encode('utf-8'))
        print('sent file transfer hint to srver')
        self.s.send(f"{fileName}{SEPARATOR}{fileSize}".encode('utf-8'))
        print('sent filename and size to server')
        with open(fileName, "rb") as f:
            while True:
                print('sending...')
                # read the bytes from the file
                bytes_read = f.read()
                if not bytes_read:
                    # file transmitting is done
                    print('done sending')
                    break
                # we use sendall to assure transimission in 
                # busy networks
                self.s.sendall(bytes_read)

    def write_msg(self):                      ##sends the message to server and server sends it to all clients
        message=f"{self.username}: {self.msg_input.get('1.0','end')}"
        self.s.send(message.encode('utf-8'))
        self.msg_input.delete('1.0','end')
    def exit(self):                                             #exit method; executed when either exit button
        # self.sock.send("exit")                      #    is pressed or 'X' is pressed
        # self.sock.send(f'{self.username}')
        self.running=False
        self.window.destroy()
        self.s.close()
        exit(0)
    def receive(self):
        while self.running:
            try:
                msg_b=self.s.recv(1024)
                msg=str(msg_b,encoding='utf-8')
                if msg=='USERNAME':                         #sends the username of the client
                    self.s.send(self.username.encode('utf-8'))
                    msg=""
                else:
                    if msg=="c":
                        toInterpret='1'
                        msg=""
                    if msg=="m":
                        toInterpret='2'
                        msg=""
                    if toInterpret=='2':
                        if self.ui_complete:                    ##if the ui is complete then it renders the msg
                            self.msgs.config(state='normal')
                            self.msgs.insert('end', msg)
                            self.msgs.yview('end')
                            self.msgs.config(state='disabled')
                        else:                                   ##if ui is incomplete then it waits for 0.5s and 
                            time.sleep(0.5)                     #renders the msg assuming ui rendering is complete
                            self.msgs.config(state='normal')
                            self.msgs.insert('end', msg)
                            self.msgs.yview('end')
                            self.msgs.config(state='disabled')
                    if toInterpret=='1':
                        if self.ui_complete:                    ##if the ui is complete then it renders the msg
                            self.users.config(state='normal')
                            self.users.replace('1.0','end',msg)
                            self.users.yview('end')
                            self.users.config(state='disabled')
                        else:                                   ##if ui is incomplete then it waits for 0.5s and 
                            time.sleep(0.5)                     #renders the msg assuming ui rendering is complete
                            self.users.config(state='normal')
                            self.users.replace('1.0','end',msg)
                            self.users.yview('end')
                            self.users.config(state='disabled')    
            except ConnectionAbortedError:
                break
            except:
                print("Unknown Error!!!")
                self.s.close()
                break


client=Client(HOST, PORT)