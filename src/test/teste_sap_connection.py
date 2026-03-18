import win32com.client

SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

print("✅ Conectado no SAP!")
print("System:", session.Info.SystemName)
print("Client:", session.Info.Client)
print("User:", session.Info.User)
print("Transaction:", session.Info.Transaction)
print("Program:", session.Info.Program, "| Screen:", session.Info.ScreenNumber)