unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.StdCtrls,
  System.Win.ComObj, Data.DB, Data.Win.ADODB, Vcl.ExtCtrls, System.Math,
  Vcl.CheckLst;

type
  TForm1 = class(TForm)
    DB: TADOConnection;
    Query: TADOQuery;
    PageControl1: TPageControl;
    FormSheet: TTabSheet;
    DeformSheet: TTabSheet;
    DeformShowButton: TButton;
    FormShowButton: TButton;
    FormListBox: TCheckListBox;
    FormLabel: TLabel;
    DeformLabel: TLabel;
    FormUnionButton: TButton;
    DeformDivideButton: TButton;
    DeformListBox: TListBox;
    procedure FormCreate(Sender: TObject);
    procedure FormShowButtonClick(Sender: TObject);
    procedure DeformShowButtonClick(Sender: TObject);
    procedure FormListBoxClickCheck(Sender: TObject);
    procedure FormListBoxDrawItem(Control: TWinControl; Index: Integer;
      Rect: TRect; State: TOwnerDrawState);
    procedure FormUnionButtonClick(Sender: TObject);
    procedure DeformDivideButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  ZenITh: OLEVariant;
  DB: TADOConnection;
  Query: TADOQuery;
  AppHWND: HWND;
  AppHandle: HWND;
  RemoteControl, IsAbort: boolean;
  CurTaskId, CurPrevFinish, CurNextStart, CurStart, CurFinish, SelMinIndex,
    SelMaxIndex: Integer;
  TaskIdList, PartIdList, CurItemIndex, ColorIdList, SelPartIdList,
    SelTaskIdList: array of Integer;

implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  try
  begin
    // подключаемся к внутренней БД
    DB.ConnectionString := ZenITh.DBParameters;
    DB.Connected := true;
    ZenITh.ExecCommand(12);
  end
  except
    on E: Exception do
      raise Exception.Create('Невозможно подключиться к внутренней БД, ошибка: '
        + E.Message);
  end;
end;

// Возвращает строку, созданную из чисел массива через запятую (для IN в запросах)
function GetStringFromArray(var Arr: array of Integer): String;
var
  i: Integer;
  Str: String;
begin
  for i := 0 to High(Arr) do
  begin
    Str := Str + IntToStr(Arr[i]);
    if i <> High(Arr) then
      Str := Str + ',';
  end;
  Result := Str;
end;

procedure TForm1.FormShowButtonClick(Sender: TObject);
var
  i, j: Integer;
  // относительно текущей выбранной операции (в Зените)
  CurPlaceId, CurGroupId, Start, Finish, MinOperationLength: Integer;
  CurPartIdList: array of Integer;
  // относительно всех кандидатов в объединение
  IsAssembly, IsValidAssembly: Boolean;
  PartIdStrList, TaskIdStrList: String;
  minStartDateTime, maxFinishDateTime: TDateTime;
label
  l1, l2;
begin
  DeformListBox.Items.Clear;
  FormUnionButton.Enabled := false;
  CurStart := 0;
  CurFinish := 0;
  FormListBox.Items.Clear;
  // определить выбранную в Зените запись
  // номер операции
  CurTaskId := ZenITh.CurrentTaskID;
  if (CurTaskId = NULL) or (CurTaskId = 0) then
  begin
    ZenITh.ShowMessage
      ('Для корректной работы надстройки необходимо выбрать операцию на Графике загрузки рабочих мест.',
      '', mb_IconStop);
    Exit;
  end;
  // номер рабочего места + вид работы + номера деталей + начало и конец выбранной операции
  // и проверка на то, чтобы GroupId совпадал
  with TADOQuery.Create(Self) do
  begin
    Connection := DB;
    SQL.Text :=
      'SELECT pt.PlaceId, pt.GroupId, pt.PartId, pt.Start, pt.Finish, pt.KindId, '
      + '(SELECT Finish FROM PlanTask p1 WHERE p1.PartId = pt.PartId AND p1.RecordPosition = pt.RecordPosition - 1) as PrevFinish, '
      + '(SELECT Start FROM PlanTask p2 WHERE p2.PartId = pt.PartId AND p2.RecordPosition = pt.RecordPosition + 1) as NextStart '
      + 'FROM PlanTask pt ' + 'WHERE pt.TaskId = ' + IntToStr(CurTaskId) + ' ';
    Open;
    First;
    i := 0;
    SetLength(CurPartIdList, RecordCount);
    CurGroupId := FieldByName('GroupId').AsInteger;
    CurPlaceId := FieldByName('PlaceId').AsInteger;
    CurStart := FieldByName('Start').AsInteger;
    CurFinish := FieldByName('Finish').AsInteger;
    CurPrevFinish := FieldByName('PrevFinish').AsInteger;
    CurNextStart := FieldByName('NextStart').AsInteger;
    while not eof do
    begin
      if FieldByName('KindId').AsInteger <> 10 then
      begin
        CurPartIdList := nil;
        FormListBox.Items.Clear;
        Close;
        Free;
        ZenITh.ShowMessage
          ('Выбранная операция не является технологической операцией.', '',
          mb_IconExclamation);
        Exit;
      end;
      // в подходящем для объединения варианте GroupId совпадает (если сборка/объединение)
      if (i > 0) and (CurGroupId <> FieldByName('GroupId').AsInteger) then
      begin
        CurPartIdList := nil;
        FormListBox.Items.Clear;
        Close;
        Free;
        ZenITh.ShowMessage
          ('Операции в выбранной сборке относятся к разным видам работ, поэтому объединение невозможно.',
          '', mb_IconStop);
        Exit;
      end;
      // ищем минимальное начало следующей операции и максимальное завершение предыдущей, чтобы отобрать соседние подходящие
      if CurPrevFinish < FieldByName('PrevFinish').AsInteger then
        CurPrevFinish := FieldByName('PrevFinish').AsInteger;
      if CurNextStart > FieldByName('NextStart').AsInteger then
        CurNextStart := FieldByName('NextStart').AsInteger;
      CurPartIdList[i] := FieldByName('PartId').AsInteger;
      Application.ProcessMessages;
      Next;
      inc(i);
    end;
    PartIdStrList := GetStringFromArray(PartIdList);
    SetLength(TaskIdList, 1);
    TaskIdList[0] := CurTaskId;

    // предыдущей и/или следующей может и не быть - тогда будут 0
    if CurPrevFinish = 0 then
      CurPrevFinish := -1;
    if CurNextStart = 0 then
      CurNextStart := -1;
    // минимальная длина операции в системе (константа)
    MinOperationLength := 0;
    SQL.Text :=
      'SELECT NumericValue FROM PlanConsts WHERE Name = ''MinOperationLength''';
    Open;
    First;
    MinOperationLength := FieldByName('NumericValue').AsInteger;

    // отбираем соседние операции, разрыв между ними должен быть меньше MinOperationLength, иначе поиск заканчиваем
    // сначала предыдущие
    IsAssembly := false;
    IsValidAssembly := true;
    Start := CurStart;
    Finish := CurFinish;
    while true do
    begin
      // +надо предусмотреть случаи, когда CurPrevFinish/CurNextStart = -1 - тогда нужно ограничение в пределах одного рабочего дня, иначе выдает безумное количество записей
      SQL.Text := 'SELECT pt.TaskId, pt.PartId, pt.KindId, ' +
        'pt.Start, pt.Finish, pt.GroupId, ' + 'pt.Assembly ' +
        'FROM PlanTask pt ' + 'WHERE pt.PlaceId = ' +
        IntToStr(CurPlaceId) + ' ' +
      // если это сборка, то у ее части может не совпадать GroupId - но нас может интересовать совпадающая часть
      // => соединять ТОЛЬКО по GroupId не выход
        'AND (pt.GroupId = ' + IntToStr(CurGroupId) + ' ' +
        'OR pt.Assembly = true) ' + 'AND pt.Finish <= ' + IntToStr(Start) +
        ' ' + // ?
      // 'OR pt.Start >= ' + IntToStr(CurFinish) + ') ' + // ?
        'ORDER BY pt.Start DESC ';
      // от самой поздней к более ранним; поскольку нужна эта сортировка, придется делать в 2 прохода, тк поиск в 2 направлениях
      Open;
      if RecordCount = 0 then
        Break;
      First;
      while not eof do
      begin
        if (Start >= MinOperationLength + FieldByName('Finish').AsInteger)
        // разрыв между операциями больше минимальной длины операции в системе -> выход из всех вложенных циклов
          or (FieldByName('KindId').AsInteger <> 10)
        // или операция не является технологической операцией
          or ((CurPrevFinish <> -1) and (FieldByName('Start').AsInteger <
          CurPrevFinish))
        // операция должна начаться не раньше, чем завершится самая поздняя из предыдущих
        then
        // хотя использование goto не рекомендуется в общем случае, в случае выхода из вложенных циклов его использовать можно
        begin
          // добавить проверку на то, что это сборка и откатываться на counter элементов назад
          Close;
          Goto l1;
        end;
        // если эта операция является частью сборки/объединения, тогда проверить целиком на GroupId
        // "горизонтальный" поиск по таблице
        if FieldByName('Assembly').AsBoolean = true then
        begin
          // надо проверить всю сборку на совпадение по GroupId как-то и отобрать те части, где он совпадет
          // и нужен флаг, который укажет, что была сборка, но мы переключились на следующую после нее запись
          IsAssembly := true;
          if FieldByName('GroupId').AsInteger <> CurGroupId then
          begin
            // если не подходит, переключаемся на следующую запись и объявляем сборку невалидной
            IsValidAssembly := false;
            Next;
            continue;
          end;
        end;
        if (FieldByName('Assembly').AsBoolean = false) and (IsAssembly = true)
        then
        begin
          IsAssembly := false;
          // и на этом остановить поиск, если сборка невалидна, или продолжить его, если сборка валидна
          if IsValidAssembly = false then
          begin
            Close;
            Goto l1;
          end;
        end;

        // если все ок, сохраняем значения атрибутов
        // в случае более ранней (относительно выбранной) операции интересно только ее начало
        if Start > FieldByName('Start').AsInteger then
          Start := FieldByName('Start').AsInteger;

        // сохраняем TaskId
        SetLength(TaskIdList, Length(TaskIdList) + 1);
        TaskIdList[High(TaskIdList)] := FieldByName('TaskId').AsInteger;
        Next;
        Application.ProcessMessages;
      end;
    end;

    // затем последующие
  l1:
    IsAssembly := false;
    IsValidAssembly := true;
    while true do
    begin
      // сначала предыдущие
      SQL.Text := 'SELECT pt.TaskId, pt.PartId, pt.KindId, ' +
        'pt.Start, pt.Finish, pt.GroupId, ' + 'pt.Assembly ' +
        'FROM PlanTask pt ' + 'WHERE pt.PlaceId = ' + IntToStr(CurPlaceId) + ' '
        + 'AND (pt.GroupId = ' + IntToStr(CurGroupId) + ' ' +
        'OR pt.Assembly = true) ' + 'AND pt.Start >= ' + IntToStr(Finish) + ' '
        + 'ORDER BY pt.Start ASC ';
      // от самой ранней к более поздним
      Open;
      if RecordCount = 0 then
        Break;
      First;
      while not eof do
      begin
        if (Finish <= FieldByName('Start').AsInteger - MinOperationLength)
        // разрыв между операциями больше минимальной длины операции в системе -> выход из всех вложенных циклов
          or (FieldByName('KindId').AsInteger <> 10)
        // или операция не является технологической операцией
          or ((CurNextStart <> -1) and (FieldByName('Finish').AsInteger >
          CurNextStart))
        // операция должна завершиться не позже, чем начнется самая ранняя из последующих
        then
        // хотя использование goto не рекомендуется в общем случае, в случае выхода из вложенных циклов его использовать можно
        begin
          Close;
          Goto l2;
        end;
        // если эта операция является частью сборки/объединения, тогда проверить целиком на GroupId
        // "горизонтальный" поиск по таблице
        if FieldByName('Assembly').AsBoolean = true then
        begin
          // надо проверить всю сборку на совпадение по GroupId как-то и отобрать те части, где он совпадет
          // и нужен флаг, который укажет, что была сборка, но мы переключились на следующую после нее запись
          IsAssembly := true;
          if FieldByName('GroupId').AsInteger <> CurGroupId then
          begin
            // если не подходит, переключаемся на следующую запись и объявляем сборку невалидной
            IsValidAssembly := false;
            Next;
          end;
        end;
        if (FieldByName('Assembly').AsBoolean = false) and (IsAssembly = true)
        then
        begin
          IsAssembly := false;
          // и на этом остановить поиск
          if IsValidAssembly = false then
          begin
            Close;
            Goto l2;
          end;
        end;

        // если все ок, сохраняем значения атрибутов
        // в случае более поздней (относительно выбранной) операции интересен только ее конец
        if Finish < FieldByName('Finish').AsInteger then
          Finish := FieldByName('Finish').AsInteger;

        // сохраняем TaskId
        SetLength(TaskIdList, Length(TaskIdList) + 1);
        TaskIdList[High(TaskIdList)] := FieldByName('TaskId').AsInteger;
        Next;
        Application.ProcessMessages;
      end;
    end;
  l2:
    TaskIdStrList := GetStringFromArray(TaskIdList);

    SQL.Text :=
      'SELECT pp.RecordPosition, pp.PartName, pt.TaskName, pt.TaskId, pt.PartId, pp.Color '
      + 'FROM PlanTask pt ' + 'INNER JOIN PlanPart pp ' +
      'ON pt.PartId = pp.PartId ' + 'WHERE pt.TaskId IN (' + TaskIdStrList +
      ') ' + 'ORDER BY Start ASC, pp.RecordPosition ASC ';
    Open;
    First;
    i := 0;
    j := 0;
    SetLength(TaskIdList, 0);
    SetLength(PartIdList, 0);
    SetLength(CurItemIndex, 0);
    SetLength(ColorIdList, 0);
    while not eof do
    begin
      // нужно перестроить TaskIdList (отсортировать по Start), чтобы соответствие в FormListBox было при выборе элементов в нем
      SetLength(TaskIdList, Length(TaskIdList) + 1);
      SetLength(PartIdList, Length(PartIdList) + 1);
      SetLength(ColorIdList, Length(ColorIdList) + 1);
      TaskIdList[i] := FieldByName('TaskId').AsInteger;
      PartIdList[i] := FieldByName('PartId').AsInteger;
      ColorIdList[i] := FieldByName('Color').AsInteger;
      // и выводим его в общий список кандидатов для объединения
      if FieldByName('TaskId').AsInteger = CurTaskId then
      begin
        FormListBox.Items.Add(FormatFloat('000', FieldByName('RecordPosition')
          .AsInteger) + ': ' + FieldByName('PartName').AsString + ' ' +
          FieldByName('TaskName').AsString);
        SetLength(CurItemIndex, Length(CurItemIndex) + 1);
        CurItemIndex[j] := i;
        SelMaxIndex := i;
        inc(j);
        if j = 1 then
        begin
          FormListBox.Selected[i] := true;
          SelMinIndex := i;
        end;
        FormListBox.Checked[FormListBox.Items.Count - 1] := true;
        // надо, чтобы в самом начале выбранная операция была отмечена (или сборка вся)
      end
      else
      begin
        FormListBox.Items.Add(FormatFloat('000', FieldByName('RecordPosition')
          .AsInteger) + ': ' + FieldByName('PartName').AsString + ' ' +
          FieldByName('TaskName').AsString);
      end;
      Next;
      inc(i);
      Application.ProcessMessages;
    end;
    TaskIdStrList := GetStringFromArray(TaskIdList);
    Close;
    Free;
  end;
end;

procedure TForm1.FormListBoxDrawItem(Control: TWinControl; Index: Integer;
  // пока выделяется только 1 главная операция в сборке, а надо все
  Rect: TRect; State: TOwnerDrawState);
var
  i: Integer;
begin
  i := 0;
  with FormListBox.Canvas do
  begin
    if (Length(CurItemIndex) > 0) and (Index = CurItemIndex[i]) then
    begin
      Font.Style := [fsBold, fsUnderline];
      inc(i);
    end;
    // Font.Color := TColor(ColorIdList[Index]);
    TextRect(Rect, Rect.Left, Rect.Top, FormListBox.Items[Index]);
  end;
end;

// обработка выбора элементов в списках + проверки
procedure TForm1.FormListBoxClickCheck(Sender: TObject);
var
  i, j, CurMinIndex, CurMaxIndex, MaxPrevFinish, MinNextStart, SelStart,
    SelFinish, RangeFinish, Counter: Integer;
  IsValidUnion: Integer;
  // -2 проблема слева, -1 проблема справа, 0 проблемы и там и там, 1 - проблем нет
label
  s1, s2;
begin
  if PageControl1.ActivePage = FormSheet then
  begin
    SetLength(SelPartIdList, 0);
    SetLength(SelTaskIdList, 0);
    IsValidUnion := 1;
    // относительно выбранной в Зените операции
    CurMinIndex := MinIntValue(CurItemIndex);
    CurMaxIndex := MaxIntValue(CurItemIndex);
    MaxPrevFinish := CurPrevFinish;
    MinNextStart := CurNextStart;
    SelStart := CurStart;
    SelFinish := CurFinish;
    if FormListBox.ItemIndex < CurMinIndex then // выбраны слева от текущей
    begin
      // случай, когда  нажимают по одной и той же ячейке
      if SelMinIndex = FormListBox.ItemIndex then
        RangeFinish := FormListBox.ItemIndex + 1
      else
        RangeFinish := FormListBox.ItemIndex;

      with TADOQuery.Create(Self) do
      begin
        Connection := DB;
        Counter := 0;
        SetLength(SelPartIdList, Length(SelPartIdList) + 1);
        SetLength(SelTaskIdList, Length(SelTaskIdList) + 1);
        SelPartIdList[0] := PartIdList[SelMaxIndex];
        SelTaskIdList[0] := TaskIdList[SelMaxIndex];
        SelMinIndex := SelMaxIndex;
        for i := SelMaxIndex downto RangeFinish do
        // идем с противоположной границы до выделенного элемента
        begin
          // проверка по PartId
          if i < SelMaxIndex then
          begin
            SetLength(SelPartIdList, Length(SelPartIdList) + 1);
            SetLength(SelTaskIdList, Length(SelTaskIdList) + 1);
            SelPartIdList[High(SelPartIdList)] := PartIdList[i];
            SelTaskIdList[High(SelTaskIdList)] := TaskIdList[i];
            if SelTaskIdList[High(SelTaskIdList)] <> SelTaskIdList
              [High(SelTaskIdList) - 1] then
            begin
              Counter := 0;
            end;
            inc(Counter);
            for j := 0 to High(SelPartIdList) - 1 do
            begin
              if SelPartIdList[j] = PartIdList[i] then
              begin
                SelMinIndex := i + Counter;
                SetLength(SelPartIdList, Length(SelPartIdList) - Counter);
                SetLength(SelTaskIdList, Length(SelTaskIdList) - Counter);
                IsValidUnion := -2;
                Goto s1;
              end;
            end;
            if i = RangeFinish then
            begin
              if ((i > 0) and (SelTaskIdList[High(SelTaskIdList)] = TaskIdList
                [i - 1])) then
              begin
                SelMinIndex := i + Counter;
                SetLength(SelPartIdList, Length(SelPartIdList) - Counter);
                SetLength(SelTaskIdList, Length(SelTaskIdList) - Counter);
                Goto s1;
              end;
            end;
          end;
          // проверка пред/след
          SQL.Text :=
            'SELECT (SELECT p1.Finish FROM PlanTask p1 WHERE p1.PartId = pt.PartId AND p1.RecordPosition = pt.RecordPosition - 1) as PrevFinish, '
            +
          // предыдущая для текущей - важен только Finish - ее может и не быть, тогда все равно
            '(SELECT p2.Start FROM PlanTask p2 WHERE p2.PartId = pt.PartId AND p2.RecordPosition = pt.RecordPosition + 1) as NextStart, '
            +
          // следующая для текущей - важен только Start - ее может и не быть, тогда все равно
            'pt.Start, pt.Finish, pt.Assembly, pt.TaskName ' +
            'FROM PlanTask pt ' + 'WHERE pt.TaskId = ' +
            IntToStr(SelTaskIdList[High(SelTaskIdList)]) + ' ' +
          // только выделенные нужны уже
            'AND pt.PartId = ' +
            IntToStr(SelPartIdList[High(SelPartIdList)]) + ' ';
          Open;
          First;
          if ((MaxPrevFinish <> -1) and (FieldByName('Start').AsInteger <
            MaxPrevFinish)) or ((FieldByName('NextStart').AsInteger < SelFinish)
            or ((SelFinish > MinNextStart) and (MinNextStart <> -1))) then
          // не CurFinish в общем случае, а SelFinish
          begin
            SelMinIndex := i + Counter;
            FormListBox.Selected[SelMinIndex] := true;
            SetLength(SelPartIdList, Length(SelPartIdList) - Counter);
            SetLength(SelTaskIdList, Length(SelTaskIdList) - Counter);
            IsValidUnion := -2;
            Goto s1;
          end
          else // проверки пройдены
          begin
            // запоминаем значение переменных
            if FieldByName('PrevFinish').AsInteger <> 0 then
            begin
              if MaxPrevFinish = -1 then
                MaxPrevFinish := FieldByName('PrevFinish').AsInteger
              else if FieldByName('PrevFinish').AsInteger > MaxPrevFinish then
                MaxPrevFinish := FieldByName('PrevFinish').AsInteger;
            end;
            if FieldByName('NextStart').AsInteger <> 0 then
            begin
              if MinNextStart = -1 then
                MinNextStart := FieldByName('NextStart').AsInteger
              else if FieldByName('NextStart').AsInteger < MinNextStart then
                MinNextStart := FieldByName('NextStart').AsInteger;
            end;
            if FieldByName('Start').AsInteger < SelStart then
              SelStart := FieldByName('Start').AsInteger;
            if FieldByName('Finish').AsInteger > SelFinish then
              SelFinish := FieldByName('Finish').AsInteger;
            SelMinIndex := i;
          end;
          Close;
        end;
        Free;
      end;
    end
    else if (FormListBox.ItemIndex >= CurMinIndex) and
      (FormListBox.ItemIndex <= CurMaxIndex) then // выбраны текущие
    begin
      SelMinIndex := CurMinIndex;
      SelMaxIndex := CurMaxIndex;
      MaxPrevFinish := CurPrevFinish;
      MinNextStart := CurNextStart;
      SelStart := CurStart;
      SelFinish := CurFinish;
      FormUnionButton.Enabled := false;
      IsValidUnion := 1;
    end
    else if FormListBox.ItemIndex > CurMaxIndex then
    // выбраны справа от текущей
    begin
      if SelMaxIndex = FormListBox.ItemIndex then
        RangeFinish := FormListBox.ItemIndex - 1
      else
        RangeFinish := FormListBox.ItemIndex;
      with TADOQuery.Create(Self) do
      begin
        Connection := DB;
        Counter := 0;
        SetLength(SelPartIdList, Length(SelPartIdList) + 1);
        SetLength(SelTaskIdList, Length(SelTaskIdList) + 1);
        SelPartIdList[0] := PartIdList[SelMinIndex];
        SelTaskIdList[0] := TaskIdList[SelMinIndex];
        SelMaxIndex := SelMinIndex;
        for i := SelMinIndex to RangeFinish do
        // идем с противоположной границы до выделенного элемента
        begin
          // проверка по PartId
          if i > SelMinIndex then
          begin
            SetLength(SelPartIdList, Length(SelPartIdList) + 1);
            SetLength(SelTaskIdList, Length(SelTaskIdList) + 1);
            SelPartIdList[High(SelPartIdList)] := PartIdList[i];
            SelTaskIdList[High(SelTaskIdList)] := TaskIdList[i];
            if SelTaskIdList[High(SelTaskIdList)] <> SelTaskIdList
              [High(SelTaskIdList) - 1] then
            begin
              Counter := 0;
            end;
            inc(Counter);
            for j := 0 to High(SelPartIdList) - 1 do
            begin
              if SelPartIdList[j] = PartIdList[i] then
              begin
                SelMaxIndex := i - Counter;
                SetLength(SelPartIdList, Length(SelPartIdList) - Counter);
                SetLength(SelTaskIdList, Length(SelTaskIdList) - Counter);
                IsValidUnion := -1;
                Goto s1;
              end;
            end;
            if i = RangeFinish then
            begin
              if ((i < High(TaskIdList)) and (SelTaskIdList[High(SelTaskIdList)
                ] = TaskIdList[i + 1])) then
              begin
                SelMaxIndex := i - Counter;
                SetLength(SelPartIdList, Length(SelPartIdList) - Counter);
                SetLength(SelTaskIdList, Length(SelTaskIdList) - Counter);
                Goto s1;
              end;
            end;
          end;
          // проверка пред/след
          SQL.Text :=
            'SELECT (SELECT p1.Finish FROM PlanTask p1 WHERE p1.PartId = pt.PartId AND p1.RecordPosition = pt.RecordPosition - 1) as PrevFinish, '
            +
          // предыдущая для текущей - важен только Finish - ее может и не быть, тогда все равно
            '(SELECT p2.Start FROM PlanTask p2 WHERE p2.PartId = pt.PartId AND p2.RecordPosition = pt.RecordPosition + 1) as NextStart, '
            +
          // следующая для текущей - важен только Start - ее может и не быть, тогда все равно
            'pt.Start, pt.Finish ' + 'FROM PlanTask pt ' + 'WHERE pt.TaskId = '
            + IntToStr(SelTaskIdList[High(SelTaskIdList)]) + ' ' +
          // только выделенные нужны уже
            'AND pt.PartId = ' +
            IntToStr(SelPartIdList[High(SelPartIdList)]) + ' ';
          Open;
          First;
          if ((MinNextStart <> -1) and (FieldByName('Finish').AsInteger >
            MinNextStart)) or ((FieldByName('PrevFinish').AsInteger > SelStart)
            or (SelStart < MaxPrevFinish)) then
          // не CurStart в общем случае, а SelStart, тк могли быть выделения ранее
          begin
            SelMaxIndex := i - Counter;
            FormListBox.Selected[SelMaxIndex] := true;
            SetLength(SelPartIdList, Length(SelPartIdList) - Counter);
            SetLength(SelTaskIdList, Length(SelTaskIdList) - Counter);
            IsValidUnion := -1;
            Goto s1;
          end
          else // проверки пройдены
          begin
            // запоминаем значение переменных
            if FieldByName('PrevFinish').AsInteger <> 0 then
            begin
              if MaxPrevFinish = -1 then
                MaxPrevFinish := FieldByName('PrevFinish').AsInteger
              else if FieldByName('PrevFinish').AsInteger > MaxPrevFinish then
                MaxPrevFinish := FieldByName('PrevFinish').AsInteger;
            end;
            if FieldByName('NextStart').AsInteger <> 0 then
            begin
              if MinNextStart = -1 then
                MinNextStart := FieldByName('NextStart').AsInteger
              else if FieldByName('NextStart').AsInteger < MinNextStart then
                MinNextStart := FieldByName('NextStart').AsInteger;
            end;
            if FieldByName('Start').AsInteger < SelStart then
              SelStart := FieldByName('Start').AsInteger;
            if FieldByName('Finish').AsInteger > SelFinish then
              SelFinish := FieldByName('Finish').AsInteger;
            SelMaxIndex := i;
          end;
          Close;
        end;
        Free;
      end;
    end;
  s1:
    // сами выделения и блокировки
    // придется идти в 2 прохода
    with TADOQuery.Create(Self) do
    begin
      Connection := DB;
      // сначала слева
      for i := SelMinIndex - 1 downto 0 do
      begin
        FormListBox.Checked[i] := false;
        FormListBox.Refresh;
        if (IsValidUnion = -2) or (IsValidUnion = 0) then
        begin
          FormListBox.ItemEnabled[i] := false;
          FormListBox.Refresh;
        end
        else
        begin
          SQL.Text := 'SELECT pt.Start ' + 'FROM PlanTask pt ' +
            'WHERE pt.TaskId = ' + IntToStr(TaskIdList[i]) + ' ' +
            'AND pt.PartId = ' + IntToStr(PartIdList[i]) + ' ';
          Open;
          First;
          if FieldByName('Start').AsInteger < MaxPrevFinish then
          begin
            FormListBox.ItemEnabled[i] := false;
            FormListBox.Refresh;
            if IsValidUnion = -1 then
              IsValidUnion := 0
            else
              IsValidUnion := -2;
          end
          else
          begin
            FormListBox.ItemEnabled[i] := true;
            FormListBox.Refresh;
          end;
          Close;
        end;
        Application.ProcessMessages;
      end;
      // затем выделенные
      for i := SelMinIndex to SelMaxIndex do
      begin
        FormListBox.Checked[i] := true;
        FormListBox.ItemEnabled[i] := true;
        Application.ProcessMessages;
        FormListBox.Refresh;
      end;
      // затем те, что справа
      for i := SelMaxIndex + 1 to High(PartIdList) do
      begin
        FormListBox.Checked[i] := false;
        FormListBox.Refresh;
        if (IsValidUnion = -1) or (IsValidUnion = 0) then
        begin
          FormListBox.ItemEnabled[i] := false;
          FormListBox.Refresh;
        end
        else
        begin
          SQL.Text := 'SELECT pt.Finish ' + 'FROM PlanTask pt ' +
            'WHERE pt.TaskId = ' + IntToStr(TaskIdList[i]) + ' ' +
            'AND pt.PartId = ' + IntToStr(PartIdList[i]) + ' ';
          Open;
          First;
          if FieldByName('Finish').AsInteger > MinNextStart then
          begin
            FormListBox.ItemEnabled[i] := false;
            FormListBox.Refresh;
            IsValidUnion := -1;
          end
          else
          begin
            FormListBox.ItemEnabled[i] := true;
            FormListBox.Refresh;
          end;
          Close;
        end;
        Application.ProcessMessages;
      end;
      Free;
      if (SelMinIndex = CurMinIndex) and (SelMaxIndex = CurMaxIndex) then
        FormUnionButton.Enabled := false
      else
        FormUnionButton.Enabled := true;
    end;
  end;
end;

// объединение операций
procedure TForm1.FormUnionButtonClick(Sender: TObject);
var
  i, MinStart, MaxFinish, SumProcessing, MinTaskId, MaxPreparation: Integer;
  minStartDateTime, maxFinishDateTime: TDateTime;
  SelTaskIdStrList, SelPartIdStrList: String;
begin
  // поместить выделенные в списке операции в строку для запросов с диапазоном
  SelTaskIdStrList := GetStringFromArray(SelTaskIdList);
  SelPartIdStrList := GetStringFromArray(SelPartIdList);
  with TADOQuery.Create(Self) do
  begin
    Connection := DB;
    SQL.Text :=
      'SELECT (SUM(p.Processing) + MAX(p.Preparation)) AS CalcProcessing, ' +
      'MIN(p.Start) AS MinStart, ' +
      'MIN(p.StartDateTime) AS MinStartDateTime, ' +
      'MIN(p.MTaskId) AS MinTaskId, ' + 'MAX(p.Preparation) AS MaxPreparation '
      + 'FROM (SELECT MIN(pt.TaskId) as MTaskId, ' +
      'MIN(pt.Processing - pt.Preparation) as Processing, ' +
      'MAX(pt.Preparation) as Preparation, ' + 'MIN(pt.TaskId) as TaskId, ' +
      'MIN(pt.Start) as Start, ' + 'MIN(pt.StartDateTime) as StartDateTime ' +
      'FROM PlanTask pt ' +
      'WHERE pt.TaskId IN (' + SelTaskIdStrList + ') ' + 'AND pt.PartId IN(' +
      SelPartIdStrList + ') ' + 'GROUP BY pt.TaskId) AS p ';

    Open;
    First; // запись будет 1, тк агрегации
    MinTaskId := FieldByName('MinTaskId').AsInteger;
    MinStart := FieldByName('MinStart').AsInteger;
    minStartDateTime := FieldByName('MinStartDateTime').AsDateTime;
    SumProcessing := FieldByName('CalcProcessing').AsInteger;
    MaxPreparation := FieldByName('MaxPreparation').AsInteger;
    Close;
    MaxFinish := MinStart + SumProcessing;
    maxFinishDateTime := ZenITh.PosToTime[MaxFinish];
    try
      DB.BeginTrans;
      ParamCheck := true;
      SQL.Text := 'UPDATE PlanTask ' + 'SET TaskId = :p1, ' + 'Start = :p2, ' +
        'Finish = :p3, ' + 'Processing = :p4, ' + 'StartDateTime = :p5, ' +
        'FinishDateTime = :p6, ' + 'Preparation = :p7, ' + 'Assembly = true ' +
        'WHERE TaskId IN(' + SelTaskIdStrList + ') ' + 'AND PartId IN( ' +
        SelPartIdStrList + ') ';
      Parameters.ParamByName('p1').Value := MinTaskId;
      Parameters.ParamByName('p2').Value := MinStart;
      Parameters.ParamByName('p3').Value := MaxFinish;
      Parameters.ParamByName('p4').Value := SumProcessing;
      Parameters.ParamByName('p5').Value := minStartDateTime;
      Parameters.ParamByName('p6').Value := maxFinishDateTime;
      Parameters.ParamByName('p7').Value := MaxPreparation;
      ExecSQL();
      DB.CommitTrans;
    except
      on E: Exception do
      begin
        DB.RollbackTrans;
        ZenITh.ShowMessage
          ('Не удалось внести данные в базу данных. Попробуйте повторить операцию',
          '', mb_IconStop);
        Close;
        Free;
        Exit;
      end;
    end;
    // обновить данные в приложении
    // задержка на 5 секунд перед обновлением
    Screen.Cursor := crAppStart;
    Sleep(5000);
    // чтобы после объединения операция в системе осталась выбранной
    ZenITh.CurrentTaskID := MinTaskId;
    // обновление данных в основной программе
    ZenITh.Refresh;
    Form1.FormShowButtonClick(FormShowButton);
    Screen.Cursor := crDefault;
    Close;
    Free;
    FormUnionButton.Enabled := false;
  end;
end;

procedure TForm1.DeformShowButtonClick(Sender: TObject);
var
  CurPlaceId, FreeTime, SumPreparation, Preparation: Integer;
begin
  DeformDivideButton.Enabled := false;
  DeformListBox.Items.Clear;
  FormListBox.Items.Clear;
  SetLength(TaskIdList, 0);
  SetLength(PartIdList, 0);
  CurTaskId := ZenITh.CurrentTaskID;
  if (CurTaskId = NULL) or (CurTaskId = 0) then
  begin
    ZenITh.ShowMessage
      ('Для корректной работы надстройки необходимо выбрать операцию на Графике загрузки рабочих мест.',
      '', mb_IconStop);
    Exit;
  end;
  SumPreparation := 0;
  Preparation := 0;
  with TADOQuery.Create(Self) do
  begin
    Connection := DB;
    SQL.Text :=
      'SELECT pp.RecordPosition, pp.PartName, pt.TaskName, pt.Assembly, ' +
      'pt.TaskId, pt.PartId, pt.PlaceId, pt.Preparation, pt.Finish, pt.Start ' +
      'FROM PlanTask pt ' + 'INNER JOIN PlanPart pp ' +
      'ON pt.PartId = pp.PartId ' + 'WHERE pt.TaskId = ' + IntToStr(CurTaskId) +
      ' ';
    Open;
    First;
    while not eof do
    begin
      if FieldByName('TaskName').AsString = '*' then
      begin
        DeformListBox.Clear;
        SetLength(TaskIdList, 0);
        SetLength(PartIdList, 0);
        ZenITh.ShowMessage
          ('Выбранную операцию нельзя разделить, поскольку она является сборочной.',
          '', mb_IconExclamation);
        // иначе не понятно, как для таких операций восстанавливать TaskName
        Exit;
      end
      else if FieldByName('Assembly').AsBoolean = false then
      begin
        DeformListBox.Clear;
        SetLength(TaskIdList, 0);
        SetLength(PartIdList, 0);
        ZenITh.ShowMessage('Выбранная операция не является объединением.', '',
          mb_IconExclamation);
        Exit;
      end
      else
      begin
        DeformListBox.Items.Add(FormatFloat('000', FieldByName('RecordPosition')
          .AsInteger) + ': ' + FieldByName('PartName').AsString + ' ' +
          FieldByName('TaskName').AsString);
        SetLength(TaskIdList, Length(TaskIdList) + 1);
        SetLength(PartIdList, Length(PartIdList) + 1);
        TaskIdList[High(TaskIdList)] := FieldByName('TaskId').AsInteger;
        PartIdList[High(PartIdList)] := FieldByName('PartId').AsInteger;
        CurPlaceId := FieldByName('PlaceId').AsInteger;
        CurStart := FieldByName('Start').AsInteger;
        CurFinish := FieldByName('Finish').AsInteger;
        Preparation := FieldByName('Preparation').AsInteger;
        SumPreparation := SumPreparation + FieldByName('Preparation').AsInteger;
        Application.ProcessMessages;
        Next;
      end;
    end;

    // проверка, что времени до следующих операций хватит на то, чтобы разделить выбранную
    SQL.Text := 'SELECT TOP 1 pt.Start ' + 'FROM PlanTask pt ' +
      'WHERE pt.PlaceId = ' + IntToStr(CurPlaceId) + ' ' + 'AND pt.Start >= ' +
      IntToStr(CurFinish) + ' ' + 'ORDER BY pt.Start ASC';
    Open;
    First;
    FreeTime := FieldByName('Start').AsInteger - CurFinish;
    if FreeTime < (SumPreparation - Preparation) then
    begin
      DeformListBox.Clear;
      SetLength(TaskIdList, 0);
      SetLength(PartIdList, 0);
      ZenITh.ShowMessage
        ('Разделение выбранной операции невозможно, поскольку не достаточно времени до следующей операции.',
        '', mb_IconExclamation);
      Exit;
    end;

    // если же все нормально, тогда можно разделять операцию
    DeformDivideButton.Enabled := true;
    Close;
    Free;
  end;
end;

procedure TForm1.DeformDivideButtonClick(Sender: TObject);
var
  NewTaskId, i, Start, Finish, SelPartId, SelTaskId: Integer;
  ProcessingList, PreparationList: array of Integer;
begin
  SetLength(ProcessingList, 0);
  SetLength(PreparationList, 0);
  with TADOQuery.Create(Self) do
  begin
    Connection := DB;
    // рассчитаем необходимые параметры - Processing и TaskId
    try
      begin
        // чтобы после объединения в системе объединенная операция осталась выбранной
        SQL.Text := 'SELECT pt.PartId AS SelPartId ' +
                    'FROM PlanTask pt ' +
                    'INNER JOIN PlanPart pp ' +
                    'ON pt.PartId = pp.PartId ' +
                    'WHERE pp.RecordPosition = (SELECT MIN(pp1.RecordPosition) ' +
                    'FROM PlanPart pp1 WHERE pp1.PartId = pp.PartId)' +
                    'AND pt.TaskId = ' + IntToStr(CurTaskId) + ' ';
        Open;
        First;
        SelPartId := FieldByName('SelPartId').AsInteger;
        // получаем новый TaskId
        NewTaskId := ZenITh.NewIDValue[208];
        for i := 0 to High(PartIdList) do
        begin
          SQL.Text :=
            'SELECT (ROUND(SWITCH(pg.portion = -1, IIF(pt.TaskVolume IS NULL, '
            + 'pp.PartVolume, pt.TaskVolume) * IIF(pt.InitProcessing IS NULL, '
            + 'pg.InitProcessing, pt.InitProcessing), ' +
            'pg.portion = 0, IIF(pt.InitProcessing IS NULL, pg.InitProcessing, pt.InitProcessing), '
            + 'pg.portion > 0,  (IIF(pt.TaskVolume IS NULL, pp.PartVolume, pt.TaskVolume) * '
            + 'IIF(pt.InitProcessing IS NULL, pg.InitProcessing, pt.InitProcessing))/pg.Portion)) '
            + '+ IIF(pt.InitPreparation IS NULL, pg.InitPreparation, pt.InitPreparation)) AS CalcProcessing, '
            + 'IIF(pt.InitPreparation IS NULL, pg.InitPreparation, pt.InitPreparation) AS CalcPreparation '
            + 'FROM  (PlanTask pt ' + 'INNER JOIN PlanPart pp ' +
            'ON (pt.PartId = pp.PartId)) ' + 'INNER JOIN PlanGrpl pg ' +
            'ON ((pt.GroupId = pg.GroupId) AND (pt.PlaceId = pg.PlaceId)) ' +
            'WHERE pt.TaskId = ' + IntToStr(CurTaskId) + ' ' +
            'AND pt.PartId = ' + IntToStr(PartIdList[i]) + ' ';
          Open;
          First;
          while not eof do
          begin
            SetLength(ProcessingList, Length(ProcessingList) + 1);
            SetLength(PreparationList, Length(PreparationList) + 1);
            ProcessingList[i] := FieldByName('CalcProcessing').AsInteger;
            PreparationList[i] := FieldByName('CalcPreparation').AsInteger;
            Next;
          end;
        end;
        for i := 0 to High(PartIdList) do
        begin
          if i = 0 then
            Start := CurStart
          else
            Start := Start + ProcessingList[i - 1];
          Finish := Start + ProcessingList[i];
          DB.BeginTrans;
          ParamCheck := true;
          if PartIdList[i] = SelPartId then
            SelTaskId := NewTaskId + i;
          SQL.Text := 'UPDATE PlanTask ' + 'SET TaskId = ' +
            IntToStr(NewTaskId + i) + ', ' + 'Assembly = false, ' +
            'Processing = ' + IntToStr(ProcessingList[i]) + ', ' +
            'Preparation = ' + IntToStr(PreparationList[i]) + ', ' + 'Start = '
            + IntToStr(Start) + ', ' + 'Finish = ' + IntToStr(Finish) + ', ' +
            'StartDateTime = :p1, ' + 'FinishDateTime = :p2 ' +
            'WHERE TaskId = ' + IntToStr(CurTaskId) + ' ' + 'AND PartId = ' +
            IntToStr(PartIdList[i]) + ' ';
          Parameters.ParamByName('p1').Value := ZenITh.PosToTime[Start];
          Parameters.ParamByName('p2').Value := ZenITh.PosToTime[Finish];
          ExecSQL;
          DB.CommitTrans;
          DeformDivideButton.Enabled := false;
        end;
      end;
    except
      begin
        DB.RollbackTrans;
        Form1.DeformShowButtonClick(DeformShowButton);
        ZenITh.ShowMessage
          ('Не удалось внести данные в базу данных. Попробуйте повторить операцию',
          '', mb_IconStop);
        Exit;
      end;
    end;
    // обновить данные в приложении
    // задержка на 5 секунд перед обновлением
    Screen.Cursor := crAppStart;
    DeformListBox.Clear;
    Sleep(5000);
    // обновление данных в основной программе
    ZenITh.Refresh;
    ZenITh.CurrentTaskID := SelTaskId;
    Screen.Cursor := crDefault;
    Close;
    Free;
  end;
end;

end.
