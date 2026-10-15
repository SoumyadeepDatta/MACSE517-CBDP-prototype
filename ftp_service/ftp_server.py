import os
import logging
from pyftpdlib.authorizers import DummyAuthorizer
from pyftpdlib.handlers import FTPHandler
from pyftpdlib.servers import FTPServer

# Enable detailed internal logging to reveal OS-level blocks
logging.basicConfig(level=logging.DEBUG)


def start_ftp():
    # Create a neutral, public directory in the current folder to avoid OS user restrictions
    inbox_dir = os.path.abspath(os.path.join(os.getcwd(), "FTP_INBOX_PUBLIC"))
    os.makedirs(inbox_dir, exist_ok=True)

    try:
        # Force the OS to grant full read/write access to this folder
        os.chmod(inbox_dir, 0o777)
    except Exception as e:
        print(f"Could not change OS folder permissions: {e}")

    authorizer = DummyAuthorizer()
    # M and T permissions added to allow file modification times to sync
    authorizer.add_user("employee_123", "password123", inbox_dir, perm="elradfmwMT")

    handler = FTPHandler
    handler.authorizer = authorizer
    handler.permit_foreign_addresses = True

    server = FTPServer(("127.0.0.1", 2123), handler)
    print(f"--- FTP SERVER RUNNING ON PORT 2123 ---")
    print(f"Inbox strictly mapped to: {inbox_dir}")
    server.serve_forever()


if __name__ == "__main__":
    start_ftp()
