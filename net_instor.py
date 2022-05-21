class creat:
    def __init__(self,id,passW) -> None:
        self.passW=self.creatpass(passW)
        self.id=id
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
        passapp=r'"\"C:/Program Files (x86)/QuOLA/quolaN0A1200075008.exe\""'
        text= f'''
        var sh = new ActiveXObject( "WScript.Shell" );
        sh.Run({passapp},1);
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
        f=open("net.js","w")
        f.write(text)
        f.close

print("学内ネットワークに自動接続ツールのインストーラーです。\nまず、idとパスワードを入力してもらいます。")
id=input("idを入力してください。＞")
passW=input("パスワードを入力してください。＞")
creat(id,passW).net()
print("net.jsを制作出来ました。")
print("net.jsをクリックしてみてください。")
input("エンターを押してください。\nプログラムが終了します。")
