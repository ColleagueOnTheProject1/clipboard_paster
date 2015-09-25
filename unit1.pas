unit Unit1;

{$mode delphi}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs,
  clipbrd, StdCtrls, Windows, Messages;

type
  { TForm1 }
  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Edit1: TEdit;
    Edit2: TEdit;
    Memo1: TMemo;
    Memo2: TMemo;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Memo2Change(Sender: TObject);
  private
    { Private declarations }
    FNextClipboardOwner: HWnd;
    function WMChangeCBChain(wParam: WParam; lParam: LParam): LRESULT;
    function WMDrawClipboard(wParam: WParam; lParam: LParam): LRESULT;
    procedure WMHotKey(var Mes: TMessage); message WM_HOTKEY;
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

var
  PrevWndProc: Windows.WNDPROC;

function WndCallback(Ahwnd: HWND; uMsg: UINT; wParam: WParam;
  lParam: LParam): LRESULT; stdcall;
begin
  if uMsg = WM_CHANGECBCHAIN then
  begin
    Result := Form1.WMChangeCBChain(wParam, lParam);
    exit;
  end
  else if uMsg = WM_DRAWCLIPBOARD then
  begin
    Result := Form1.WMDrawClipboard(wParam, lParam);
    exit;
  end;
  Result := CallWindowProc(PrevWndProc, Ahwnd, uMsg, WParam, LParam);
end;
procedure TForm1.WMHotKey(var Mes: TMessage);
var i,j,k,l:ShortInt;
procedure Past();
begin
   keybd_event(VK_CONTROL, 0, 0, 0); //Нажатие control
   keybd_event(ord('V'), 0, 0, 0); //Нажатие V
   keybd_event(ord('V'), 0, KEYEVENTF_KEYUP, 0); //Отпустить V
   keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0); //Отпустить controle
   sleep(100);
end;
procedure Clear();
begin
  keybd_event(VK_HOME, 0, 0, 0); //Нажатие home
  keybd_event(VK_HOME, 0, KEYEVENTF_KEYUP, 0); //Отпустить home
  keybd_event(VK_SHIFT, 0, KEYEVENTF_EXTENDEDKEY or KEYEVENTF_KEYUP, 0);
  keybd_event(VK_SHIFT, 0, KEYEVENTF_EXTENDEDKEY, 0);//Нажатие shift
  keybd_event(VK_END, 0, 0, 0); //Нажатие END
  keybd_event(VK_END, 0, KEYEVENTF_KEYUP, 0); //Отпустить END
  keybd_event(VK_SHIFT, 0, KEYEVENTF_EXTENDEDKEY or KEYEVENTF_KEYUP, 0);//отпустить shift
  keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0);
end;

begin
  keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);//отпустить контрол
  keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);//отпустить альт
  k := StrToInt(Edit1.Text);
  l := StrToInt(Edit2.Text);
  for i:= 0 to Memo1.Lines.Count - 1 do
  begin
     Clipboard.AsText := Memo1.Lines[i];
     Clear();
     Past;
     for j:=1 to k do
     begin
       keybd_event(VK_TAB, 0, KEYEVENTF_EXTENDEDKEY, 0); //Нажатие tab
       keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0); //Отпустить tab
     end;
     if((i+1) mod l = 0)then
     begin
       keybd_event(VK_TAB, 0, KEYEVENTF_EXTENDEDKEY, 0); //Нажатие tab
       keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0); //Отпустить tab
     end;
  end;
end;
procedure TForm1.FormCreate(Sender: TObject);
var  myHotKey:AnsiChar = 'E';
begin
  PrevWndProc := Windows.WNDPROC(
    SetWindowLong(Self.Handle, GWL_WNDPROC, PtrInt(@WndCallback)));
  FNextClipboardOwner := SetClipboardViewer(Self.Handle);
  RegisterHotKey(Form1.Handle, 1 ,MOD_ALT, Ord(myHotKey));
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  Edit1.Text := IntToStr(StrToInt(Edit1.Text) + 1);
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  Edit2.Text := IntToStr(StrToInt(Edit2.Text) - 1);
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
  Edit2.Text := IntToStr(StrToInt(Edit2.Text) + 1);
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  Edit1.Text := IntToStr(StrToInt(Edit1.Text) - 1);
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  ChangeClipboardChain(Handle, FNextClipboardOwner);
end;

procedure TForm1.Memo2Change(Sender: TObject);
begin

end;

function TForm1.WMChangeCBChain(wParam: WParam; lParam: LParam): LRESULT;
var
  Remove, Next: THandle;
begin
  Remove := WParam;
  Next := LParam;
  if FNextClipboardOwner = Remove then
    FNextClipboardOwner := Next
  else if FNextClipboardOwner <> 0 then
    SendMessage(FNextClipboardOwner, WM_ChangeCBChain, Remove, Next);
end;

function TForm1.WMDrawClipboard(wParam: WParam; lParam: LParam): LRESULT;
var s:String;
    i:byte;
begin
  if Clipboard.HasFormat(CF_TEXT) then
  begin
    s :=  Clipboard.AsText;
    s := StringReplace(s, ' ', #13#10, [rfReplaceAll, rfIgnoreCase]);
    s := StringReplace(s, ',' + #13#10, ' ', [rfReplaceAll, rfIgnoreCase]);
    for i:=0 to Memo2.Lines.Count - 1 do
    begin
        s := StringReplace(s, #13#10 + Memo2.Lines[i], ' ' + Memo2.Lines[i], [rfReplaceAll]);
    end;
    Memo1.Text := s;
  end;
  SendMessage(FNextClipboardOwner, WM_DRAWCLIPBOARD, 0, 0);
  Result := 0;
end;

end.
