var sh = new ActiveXObject( "WScript.Shell" );
sh.Run("\"C:/Program Files (x86)/QuOLA/quolaN0A1200075008.exe\"",1);
WScript.Sleep( 3000 );
id=""
pass=""
sh.SendKeys(id);WScript.Sleep( 500 );
sh.SendKeys("{TAB}");WScript.Sleep( 500 );
sh.SendKeys(pass);WScript.Sleep( 1000 );
sh.SendKeys("{ENTER}");WScript.Sleep( 1000 );
sh.SendKeys("{TAB}");WScript.Sleep( 500 );
sh.SendKeys("{ENTER}");WScript.Sleep( 2000 );
sh.SendKeys("^w");WScript.Sleep( 2000 );
sh = null;