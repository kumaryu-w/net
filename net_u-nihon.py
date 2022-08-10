import os
import shutil
from time import sleep
class creat:
    def __init__(self,id,passW,app_pass) -> None:
        self.passW=self.creatpass(passW)
        self.id=id
        self.app_pass=app_pass
    def creatpass(self,passW):
        i=0
        lists=list(passW)
        listpass=lists
        for li in lists:
            for j in ["+","^","%","[","]","{","}","(",")","~"]:
                if li==j:
                    listpass[i]="{"+str(j)+"}"
            i+=1
        return "".join(listpass)
        
    def net(self):
        text= f'''
        var sh = new ActiveXObject( "WScript.Shell" );
        sh.Run("{self.app_pass}",1);
        WScript.Sleep( 3000 );
        id="{self.id}"
        pass="{self.passW}"
        sh.SendKeys(id);WScript.Sleep( 500 );
        sh.SendKeys("{{TAB}}");WScript.Sleep( 500 );
        sh.SendKeys(pass);WScript.Sleep( 1000 );
        sh.SendKeys("{{ENTER}}");WScript.Sleep( 1000 );
        sh.SendKeys("{{TAB}}");WScript.Sleep( 500 );
        sh.SendKeys("{{ENTER}}");WScript.Sleep( 2000 );
        sh.SendKeys("^w");WScript.Sleep( 2000 );
        sh = null;'''
        print(text)
        f=open("net_u-nihon.js","w")
        f.write(text)
        f.close

print("学内ネットワークに自動接続ツールのインストーラーです。\n")

while(True):    
    print("Login.exeを保存するディレクトリーを生成します。\n(初期設定　C:/net_u-nihon/Login.exe)\n")
    sleep(1)
    selct=input("初期設定のままで[y]\nディレクトリーの詳細設定がしたい[n]\n[y]or[n]?")
    if selct == "y":
        cr_dr_pass="C:/net_u-nihon"
    elif selct == "n":
        cr_dr_pass=input("保存したい場所の絶対パスを設定してください。>>")
        
    try:
        os.makedirs(cr_dr_pass)
    except FileExistsError:
        break
    except:    
        print("正しい絶対パスを入力してください\n")
        continue
    break
try:
    app_pass=cr_dr_pass+"/Login.exe"
    shutil.copy2("Login.exe",app_pass)
except PermissionError:
    print("ファイルがすでにあるようです。プロセスをスキップします。")


print("次にログインidとパスワードを設定してください。")
while(True):
    id=input("idを入力してください。＞")
    passW=input("パスワードを入力してください。＞")
    ch=input("もう一度入力したい[y]\いいえ[n]\n[y]or[n]?")
    if ch != "y":
        break
    else:print("しゃーない！もっかい入力させてやんよ！")
    
print("net_u-nihon.jsを生成します。")
creat(id,passW,app_pass).net()
print("net_u-nihon.jsの生成が出来ました。")
sleep(0.5)
print("ファイルをデスクトップへ移動します。")

desktop_path = os.getenv("HOMEDRIVE") + os.getenv("HOMEPATH") + "\\Desktop"
shutil.move("net_u-nihon.js",desktop_path+"\\net_u-nihon.js")
print("成功")
input("エンターを押してください。\nプログラムが終了します。")

