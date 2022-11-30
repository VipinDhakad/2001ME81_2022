import socket
import threading
import time

HOST='127.0.0.1'
PORT=9090

server=socket.socket(socket.AF_INET,socket.SOCK_STREAM)
server.bind((HOST,PORT))

server.listen()

clients=[]
nicknames=[]


def broadcast(message):          #sending the content of message to all the connected clients
    for client in clients:
        client.send(message)

def handle(client):
    while True:
        try:
            message=client.recv(1024)
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