"""原始 TCP 转储服务器：打印收到的字节（诊断 vite 代理转发的畸形请求）。"""
import socket
import sys

HOST, PORT = "127.0.0.1", 8100

srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
srv.bind((HOST, PORT))
srv.listen(5)
print(f"listening on {HOST}:{PORT}", flush=True)

while True:
    conn, addr = srv.accept()
    print(f"\n=== connection from {addr} ===", flush=True)
    try:
        conn.settimeout(3)
        while True:
            try:
                data = conn.recv(65536)
            except socket.timeout:
                break
            if not data:
                break
            print(f"--- {len(data)} bytes ---", flush=True)
            print(repr(data), flush=True)
            # 回一个合法响应，便于客户端侧观察
            conn.sendall(b"HTTP/1.1 200 OK\r\nContent-Length: 2\r\nConnection: close\r\n\r\nok")
            break
    except Exception as exc:
        print(f"error: {exc!r}", flush=True)
    finally:
        conn.close()
