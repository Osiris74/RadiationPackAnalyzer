unit Serial_Port;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ComObj, ExtCtrls, Buttons, Menus, LazSerial, FileUtil,
  StrUtils, EpikTimer;


type
  TWMCopyData = packed record
    Msg: Cardinal;
    From: HWND;
    CopyDataStruct: PCopyDataStruct;
    Result: Longint;
  end;


const
  cmRxByte = wm_User+$55;
  WM_COPYDATA = wm_User+$3000;
  ALLBYTES = 15360;

type

  { TForm1 }

  TForm1 = class(TForm)
    EpikTimer1: TEpikTimer;
    MainMenu1: TMainMenu;
    FileMenuItem: TMenuItem;
    CreateLogMenuItem: TMenuItem;
    OpenDialog1: TOpenDialog;
    OpenLogMenuItem: TMenuItem;
    SaveDialog1: TSaveDialog;
    StatusBar1: TStatusBar;
    SpeedBox: TComboBox;
    PortBox: TComboBox;
    Label1: TLabel;
    Button1: TButton;
    Label2: TLabel;
    DataGroup: TRadioGroup;
    StopGroup: TRadioGroup;
    ParityBox: TComboBox;
    Label3: TLabel;
    SpeedButton1: TSpeedButton;
    RxMem: TMemo;
    SpeedButton3: TSpeedButton;
    Timer1: TTimer;
    SystemSleepTimer: TTimer;
    procedure CreateLogMenuItemClick(Sender: TObject);
    procedure OpenLogMenuItemClick(Sender: TObject);
    procedure RecivBytes(var Msg : TMessage); message cmRxByte;
    procedure PortBoxEnter(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SystemSleepTimerTimer(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure WMCopyData(var Msg: TWMCopyData); message WM_COPYDATA;
    procedure SpeedButton3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmMain: TForm1;
  autoscroll:Bool;
  isStopped: Bool;
  //Probably it is unnecesary here
  byteCounter : Long;
  wrongBytes : integer;
  globalWrongBytes : Long;
  LogFileName : String;
  firstByte : Bool;
  LogFileNameCounter : Integer;

implementation
uses MyComm, SystemSleepUnit;
{$R *.dfm}


//===================================RecivBytes===================================


procedure TForm1.WMCopyData(var Msg: TWMCopyData);
var
  sText: array[0..49] of Char;
  i:integer;
  tmp:String;

  s: PChar;
begin
  // generate text from parameter
 // anzuzeigenden Text aus den Parametern generieren


 s := PChar(Msg.CopyDataStruct.lpData) ;

 RxMem.Lines.Add(s);
  {
 StrLCopy(sText, Msg.CopyDataStruct.lpData, Msg.CopyDataStruct.cbData);
  // write received text
  for i:=0 to 49 do begin
   tmp:=tmp+sText[i];
  end;
 RxMem.Lines.Add(tmp);
 end;
 }
end;

      procedure TForm1.RecivBytes(var Msg: TMessage);   //Messages from second thread
       var
        s:string;
        message : PMyMsg;
          begin
                      if firstByte = true then EpikTimer1.Start;

                      firstByte := false;

                      message := Pointer(Msg.LParam);
                      s := message.s;
                      byteCounter := byteCounter + message.bytesCounter;
                      wrongBytes := message.wrongBytes;
                      globalWrongBytes := globalWrongBytes + wrongBytes;
                      RxMem.Lines.Add(string(s));
                      RxMem.Lines.Add(IntToStr(byteCounter) + '___' + 'Error:' + FloatToStrF((wrongBytes/message.bytesCounter)*100, ffFixed, 0, 2) + ' %');
                      RxMem.Lines.Add('');

                      Timer1.Enabled := False;
                      Timer1.Enabled := True;

                      Dispose(message);
          end;

procedure TForm1.CreateLogMenuItemClick(Sender: TObject);
var
 txtFile : TextFile;
 tmp : String;
begin

     if LogFileNameCounter = 0 then
     begin
      SaveDialog1.Execute;
      LogFileName := SaveDialog1.FileName;
      assignFile(txtFile, LogFileName);
      rewrite(txtFile);
      closefile(txtFile);
      LogFileNameCounter := LogFileNameCounter + 1;
     end

     else
     begin

      if Pos('(', LogFileName) <> 0 then
      begin
         Delete(LogFileName, Pos('(', LogFileName), Pos('.txt', LogFileName) - Pos('(', LogFileName));
      end;
           tmp := '(' + IntToStr(LogFileNameCounter) + ')';
           Insert(tmp, LogFileName, Pos('.txt', LogFileName));

      tmp := Copy(LogFileName, LastDelimiter('\', LogFileName) + 1, Pos('.txt', LogFileName) - LastDelimiter('\', LogFileName) - 1);

      SaveDialog1.FileName := tmp;
      SaveDialog1.Execute;
      LogFileName := SaveDialog1.FileName;
      assignFile(txtFile, LogFileName);
      rewrite(txtFile);
      closefile(txtFile);
      LogFileNameCounter := LogFileNameCounter + 1;
     end;
end;

procedure TForm1.OpenLogMenuItemClick(Sender: TObject);
begin
     OpenDialog1.FileName:=LogFileName;
     OpenDialog1.Execute;
end;


//===================================RecivBytes===================================

//********************************************************************************

//==================================FillCommList==================================
   //Процедура поиска свободных COM-портов
  function FillCommList( List : Tstrings ): integer;
      var
        ComName: string;
        i: integer;
        pPath : pchar;
        Size : Cardinal;
      begin
        List.Clear();               //Очистка предыдущих значений
        pPath := Stralloc( 256 );   //Задание размера буфера
          try                       //Пробуем найти порты
            for i := 1 to 99 do
              begin
                pPath[0] := #0;
                ComName := 'COM' + inttostr( i );
                QueryDosDevice( pchar( ComName), pPath, Size );  //Спец функция
                                                                 //поиска подключенных
                                                                 //устройств
                if CompareText( pPath, '' ) <> 0 then            //Добавление в список
                List.AddObject( ComName, pointer( i ) );
            end;
          finally
            strdispose( pPath );                                 //Процедура освобождения строки
          end;
        result := List.Count;
      end;

//==================================FillCommList==================================

//********************************************************************************

//=================================PortBoxEnter===================================

    //Заполнение СОМ-портоа после нажатия на BOX
    procedure TForm1.PortBoxEnter(Sender: TObject);
      begin
        FillCommList(PortBox.Items);
      end;
//=================================PortBoxEnter===================================

//********************************************************************************

//==================================FormCreate====================================

    procedure TForm1.FormCreate(Sender: TObject);
      begin
          DataGroup.ItemIndex := 3;
          StopGroup.ItemIndex := 1;
          ParityBox.ItemIndex := 0;
          //SpeedBox.ItemIndex := 8;
          LogFileNameCounter := 0;
          SystemSleepTimer.Enabled:=true;
      end;
//==================================FormCreate====================================

//********************************************************************************

//=================================ButtonsClick===================================

    procedure TForm1.Button1Click(Sender: TObject);
      begin
          RxMem.Lines.Clear;
      end;


    procedure TForm1.SpeedButton1Click(Sender: TObject);
    var
      stopbits:integer;
      begin
            if StartService = true then ShowMessage('Com-Port Started Succesfully');
            isStopped:=false;

      end;

    procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
      begin
          if isStopped=false then
          StopService;
      end;

procedure TForm1.SystemSleepTimerTimer(Sender: TObject);
begin
   SystemSleepTimer.Enabled := false;
   SetThreadExecutionState(ES_CONTINUOUS or ES_SYSTEM_REQUIRED or ES_DISPLAY_REQUIRED);
   SystemSleepTimer.Enabled := true;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
var
  txtFile : TextFile;
  data : String;
  tmp : String;
begin

    Timer1.Enabled:= false;

    if not FileExists(LogFileName) then begin
       CreateLogMenuItemClick(Sender);
    end;

    AssignFile(txtFile, LogFileName);
    Append(txtFile);

    tmp := EpikTimer1.Elapsed.ToString;
    Delete(tmp, 0, 5);
    data := 'Bytes received: ' + IntToStr(byteCounter) + '; Bytes lost: ' + IntToStr(ALLBYTES - byteCounter) +  '; Wrong bytes: ' + FloatToStrF((globalWrongBytes/byteCounter)*100, ffFixed, 0, 2) + ' %' +'; Timing: ' + tmp;

    writeln(txtFile, data);
    CloseFile(txtFile);

    StatusBar1.SimpleText := data;
    byteCounter := 0;
    globalWrongBytes := 0;
    firstByte := true;
    EpikTimer1.Clear();

end;


procedure TForm1.SpeedButton3Click(Sender: TObject);
begin
     StopService;
     ShowMessage('Com-Port Stopped Succesfully');
     isStopped:=true;
end;

end.
