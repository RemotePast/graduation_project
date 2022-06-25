library CombTechOperations;

uses
  System.SysUtils,
  Classes,
  Windows,
  Vcl.Forms,
  Vcl.Controls,
  ComObj,
  Variants,
  ADODB,
  Main in 'Main.pas' {Form1};

{$R *.res}

procedure AfterLoad(AppHandle: HWND); stdcall;
begin
  try
  Application.Initialize;
    try
      ZenITh := CreateOleObject('ZenPlan.ZenApp');
    except
      Application.MessageBox('Невозможно запустить сервер Zenith SPPS.' +
                             'Для регистрации OLE-сервера запустите ' +
                             'файл ZenPlan.exe c параметром /regserver',
                             'Действие невозможно!', mb_IconStop);
      Exit; {!!}
    end;
    AppHWND := Application.Handle;
    Application.MainFormOnTaskbar := True;
    Application.CreateForm(TForm1, Form1);
  Application.Run;
  finally
    Application.Handle := AppHWND;
    ZenITh := Unassigned;
  end;
end;

exports
AfterLoad;

begin
end.
