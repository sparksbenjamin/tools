import paramiko
import time
from typing import Dict, Any
uname = "<URSER>"
p = "<PASSWORD>"

def test_login(host: str, username: str, password: str,
               port: int = 22, timeout: float = 5.0) -> Dict[str, Any]:
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        client.connect(hostname=host, port=port, username=username, password=password,
                       timeout=timeout, banner_timeout=timeout, auth_timeout=timeout,
                       allow_agent=False, look_for_keys=False)
        client.close()
        #return {"success": True}
        return "True"
    except paramiko.AuthenticationException:
        #return {"success": False, "reason": "authentication_failed"}
        return "False, Bad username or password"
    except Exception as e:
        #return {"success": False, "reason": str(e)}
        return f"False, {str(e)}"
    finally:
        try:
            client.close()
        except Exception:
            pass
seen = set()
failed = 0
with open("targets.txt") as file:
    for line in file.readlines():
        hname = line.strip().replace('"','')
        if hname != "" and hname not in seen:
            seen.add(hname)
            login = test_login(hname,uname,p)
            if "False" in login:
                failed =+ 1
            print(f'{hname},{login}')
            #Avoild Locking the account. 
            if failed > 2:
                failed = 0
                time.sleep(30)
            
