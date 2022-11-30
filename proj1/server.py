import socket
import threading
import time
import os

HOST='127.0.0.1'
PORT=9090

server=socket.socket(socket.AF_INET,socket.SOCK_STREAM)
server.bind((HOST,PORT))

server.listen()

BUFFER_SIZE = 4096
SEPARATOR = "<SEPARATOR>"

clients=[]
nicknames=[]


def broadcast(message):          #sending the content of message to all the connected clients
    for client in clients:
        client.send(message)

def handle(client):
    while True:
        try:
            message=client.recv(1024)
            if message.decode('utf-8')=='!FTP':
                print('from SERVER: received !FTP')
                received = client.recv(1024).decode('utf-8')
                fileName, fileSize = received.split(SEPARATOR)
                print(f'received {fileName} and {fileSize}')
                # remove absolute path if there is
                fileName = os.path.basename(fileName)
                # convert to integer
                fileSize = int(fileSize)
                with open(fileName, "wb") as f:
                    while True:
                        # read 1024 bytes from the socket (receive)
                        bytes_read = client.recv(BUFFER_SIZE)
                        if not bytes_read:    
                            # nothing is received
                            # file transmitting is done
                            # broadcast(f)
                            break
                        # write to the file the bytes we just received
                        print('wrote into file...')
                        f.write(bytes_read)
                        broadcast(bytes_read)
                print("Done receiving the file\n")
                # broadcast("m".encode('utf-8'))
                # broadcast(f"{fileName} Received!")
            else:
                broadcast("m".encode('utf-8'))
                broadcast(message)
        except:
            index=clients.index(client)
            clients.remove(client)
            client.close()
            nickname=nicknames[index]
            nicknames.remove(nickname)
            break

def receive():
    while True:
        client,address=server.accept()
        
        client.send("USERNAME".encode('utf-8'))
        nickname=client.recv(1024).decode('utf-8')

        nicknames.append(nickname)
        names = '\n'.join([str(elem) for i,elem in enumerate(nicknames)])
        clients.append(client)
        broadcast("c".encode('utf-8'))
        time.sleep(0.1)
        broadcast(f'{names}\n'.encode('utf-8'))
        thread=threading.Thread(target=handle,args=(client,))
        thread.start()
        # broadcast("m".encode('utf-8'))
        # time.sleep(0.01)
        # broadcast(f"{nickname} Connected\n".encode('utf-8'))

print("**********Server Started**********")
receive()