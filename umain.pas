unit umain;

{ Scanlog Importer
  Import CSV scanlog mesin absensi (PIN, NIP, Nama, Jabatan, Departemen,
  Kantor, Tanggal, Scan 1, Scan 2, Scan 3) -> tampil di grid, cari/filter,
  rekap per pegawai, export hasil filter.

  Tanpa DLL tambahan (memakai TBufDataset). Cocok untuk Windows 7. }

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ExtCtrls,
  ComCtrls, DBGrids, db, BufDataset;

type
  { TfrmMain }
  TfrmMain = class(TForm)
  private
    pnlTop: TPanel;
    btnOpen: TButton;
    btnExport: TButton;
    lblCari: TLabel;
    edtCari: TEdit;
    pc: TPageControl;
    tsData: TTabSheet;
    tsRekap: TTabSheet;
    gridData: TDBGrid;
    gridRekap: TDBGrid;
    dsData: TDataSource;
    dsRekap: TDataSource;
    bufData: TBufDataset;
    bufRekap: TBufDataset;
    sb: TStatusBar;
    dlgOpen: TOpenDialog;
    dlgSave: TSaveDialog;
    FFileName: string;
    procedure BuildUI;
    procedure btnOpenClick(Sender: TObject);
    procedure btnExportClick(Sender: TObject);
    procedure edtCariChange(Sender: TObject);
    procedure bufDataFilterRecord(DataSet: TDataSet; var Accept: Boolean);
    procedure PrepareDataset;
    procedure PrepareRekap;
    procedure LoadCSV(const AFileName: string);
    procedure BuildRekap;
    procedure FormatGrids;
    procedure UpdateStatus;
    procedure FormShowHandler(Sender: TObject);
  public
    constructor Create(TheOwner: TComponent); override;
    destructor Destroy; override;
  end;

var
  frmMain: TfrmMain;

implementation

type
  TRekap = class
    PIN, Nama: string;
    Hari, HariLengkap, Menit: Integer;
  end;

{ ---------- util ---------- }

function DetectDelimiter(const Line: string): Char;
var
  i, cComma, cSemi, cTab: Integer;
  inQ: Boolean;
begin
  cComma := 0; cSemi := 0; cTab := 0; inQ := False;
  for i := 1 to Length(Line) do
    case Line[i] of
      '"': inQ := not inQ;
      ',': if not inQ then Inc(cComma);
      ';': if not inQ then Inc(cSemi);
      #9:  if not inQ then Inc(cTab);
    end;
  Result := ',';
  if (cSemi > cComma) and (cSemi >= cTab) then Result := ';'
  else if (cTab > cComma) and (cTab > cSemi) then Result := #9;
end;

procedure SplitCSV(const Line: string; Delim: Char; Fields: TStrings);
var
  i: Integer;
  cur: string;
  inQ: Boolean;
begin
  Fields.Clear;
  cur := '';
  inQ := False;
  i := 1;
  while i <= Length(Line) do
  begin
    if inQ then
    begin
      if Line[i] = '"' then
      begin
        if (i < Length(Line)) and (Line[i + 1] = '"') then
        begin
          cur := cur + '"';
          Inc(i);
        end
        else
          inQ := False;
      end
      else
        cur := cur + Line[i];
    end
    else
    begin
      if Line[i] = '"' then
        inQ := True
      else if Line[i] = Delim then
      begin
        Fields.Add(cur);
        cur := '';
      end
      else
        cur := cur + Line[i];
    end;
    Inc(i);
  end;
  Fields.Add(cur);
end;

function CsvQuote(const S: string; Delim: Char): string;
begin
  if (Pos(Delim, S) > 0) or (Pos('"', S) > 0) or (Pos(#10, S) > 0) then
    Result := '"' + StringReplace(S, '"', '""', [rfReplaceAll]) + '"'
  else
    Result := S;
end;

{ dd-mm-yyyy, dd/mm/yyyy atau yyyy-mm-dd }
function ParseTanggal(const S: string; out D: TDateTime): Boolean;
var
  p: TStringList;
  a, b, c: Integer;
  t: string;
begin
  Result := False;
  t := StringReplace(Trim(S), '/', '-', [rfReplaceAll]);
  p := TStringList.Create;
  try
    p.Delimiter := '-';
    p.StrictDelimiter := True;
    p.DelimitedText := t;
    if p.Count <> 3 then Exit;
    a := StrToIntDef(p[0], -1);
    b := StrToIntDef(p[1], -1);
    c := StrToIntDef(p[2], -1);
    if (a < 0) or (b < 0) or (c < 0) then Exit;
    if Length(Trim(p[0])) = 4 then
      Result := TryEncodeDate(a, b, c, D)      // yyyy-mm-dd
    else
    begin
      if c < 100 then c := c + 2000;
      Result := TryEncodeDate(c, b, a, D);     // dd-mm-yyyy
    end;
  finally
    p.Free;
  end;
end;

{ hh:mm[:ss] -> menit sejak 00:00 (detik diabaikan), -1 bila kosong/invalid }
function ParseJamMenit(const S: string): Integer;
var
  t: string;
  p: Integer;
  h, m: Integer;
begin
  Result := -1;
  t := Trim(S);
  if t = '' then Exit;
  p := Pos(':', t);
  if p = 0 then Exit;
  h := StrToIntDef(Copy(t, 1, p - 1), -1);
  Delete(t, 1, p);
  p := Pos(':', t);
  if p > 0 then t := Copy(t, 1, p - 1);
  m := StrToIntDef(t, -1);
  if (h < 0) or (h > 23) or (m < 0) or (m > 59) then Exit;
  Result := h * 60 + m;
end;

function MenitToStr(AMenit: Integer): string;
begin
  Result := Format('%d:%.2d', [AMenit div 60, AMenit mod 60]);
end;

function FindCol(Header: TStrings; const Name: string): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to Header.Count - 1 do
    if SameText(Trim(Header[i]), Name) then
      Exit(i);
end;

function GetCol(Row: TStrings; Idx: Integer): string;
begin
  if (Idx >= 0) and (Idx < Row.Count) then
    Result := Trim(Row[Idx])
  else
    Result := '';
end;

{ ---------- TfrmMain ---------- }

constructor TfrmMain.Create(TheOwner: TComponent);
begin
  inherited CreateNew(TheOwner, 0);
  Caption := 'Scanlog Importer - Import CSV Absensi';
  Width := 1000;
  Height := 600;
  Position := poScreenCenter;
  BuildUI;
  PrepareDataset;
  PrepareRekap;
  UpdateStatus;
  OnShow := @FormShowHandler;
end;

{ Parameter: scanlog.exe [file.csv] [rekap]
  Dipakai untuk membuka CSV langsung saat start (dan membuat screenshot otomatis). }
procedure TfrmMain.FormShowHandler(Sender: TObject);
begin
  OnShow := nil;
  if (ParamCount >= 1) and FileExists(ParamStr(1)) then
  begin
    try
      LoadCSV(ParamStr(1));
    except
      // abaikan; pengguna bisa membuka manual
    end;
    if SameText(ParamStr(2), 'rekap') then
      pc.ActivePage := tsRekap;
  end;
end;

destructor TfrmMain.Destroy;
begin
  bufData.Close;
  bufRekap.Close;
  inherited Destroy;
end;

procedure TfrmMain.BuildUI;
begin
  dlgOpen := TOpenDialog.Create(Self);
  dlgOpen.Title := 'Pilih file CSV scanlog';
  dlgOpen.Filter := 'File CSV (*.csv)|*.csv|Semua file (*.*)|*.*';
  dlgOpen.Options := dlgOpen.Options + [ofFileMustExist];

  dlgSave := TSaveDialog.Create(Self);
  dlgSave.Title := 'Export hasil ke CSV';
  dlgSave.Filter := 'File CSV (*.csv)|*.csv';
  dlgSave.DefaultExt := 'csv';
  dlgSave.Options := dlgSave.Options + [ofOverwritePrompt];

  pnlTop := TPanel.Create(Self);
  pnlTop.Parent := Self;
  pnlTop.Align := alTop;
  pnlTop.Height := 44;
  pnlTop.BevelOuter := bvNone;

  btnOpen := TButton.Create(Self);
  btnOpen.Parent := pnlTop;
  btnOpen.SetBounds(8, 8, 110, 28);
  btnOpen.Caption := 'Buka CSV...';
  btnOpen.OnClick := @btnOpenClick;

  btnExport := TButton.Create(Self);
  btnExport.Parent := pnlTop;
  btnExport.SetBounds(124, 8, 130, 28);
  btnExport.Caption := 'Export Hasil...';
  btnExport.OnClick := @btnExportClick;

  lblCari := TLabel.Create(Self);
  lblCari.Parent := pnlTop;
  lblCari.SetBounds(272, 14, 90, 16);
  lblCari.Caption := 'Cari PIN/Nama:';

  edtCari := TEdit.Create(Self);
  edtCari.Parent := pnlTop;
  edtCari.SetBounds(370, 10, 220, 24);
  edtCari.OnChange := @edtCariChange;

  sb := TStatusBar.Create(Self);
  sb.Parent := Self;
  sb.Align := alBottom;
  sb.SimplePanel := True;

  pc := TPageControl.Create(Self);
  pc.Parent := Self;
  pc.Align := alClient;

  tsData := TTabSheet.Create(pc);
  tsData.PageControl := pc;
  tsData.Caption := 'Data Scanlog';

  tsRekap := TTabSheet.Create(pc);
  tsRekap.PageControl := pc;
  tsRekap.Caption := 'Rekap per Pegawai';

  bufData := TBufDataset.Create(Self);
  bufData.OnFilterRecord := @bufDataFilterRecord;
  bufRekap := TBufDataset.Create(Self);

  dsData := TDataSource.Create(Self);
  dsData.DataSet := bufData;
  dsRekap := TDataSource.Create(Self);
  dsRekap.DataSet := bufRekap;

  gridData := TDBGrid.Create(Self);
  gridData.Parent := tsData;
  gridData.Align := alClient;
  gridData.DataSource := dsData;
  gridData.ReadOnly := True;
  gridData.Options := gridData.Options - [dgEditing] + [dgRowSelect];

  gridRekap := TDBGrid.Create(Self);
  gridRekap.Parent := tsRekap;
  gridRekap.Align := alClient;
  gridRekap.DataSource := dsRekap;
  gridRekap.ReadOnly := True;
  gridRekap.Options := gridRekap.Options - [dgEditing] + [dgRowSelect];
end;

procedure TfrmMain.PrepareDataset;
begin
  bufData.Close;
  bufData.Filtered := False;
  bufData.FieldDefs.Clear;
  bufData.FieldDefs.Add('PIN', ftString, 20);
  bufData.FieldDefs.Add('NIP', ftString, 30);
  bufData.FieldDefs.Add('Nama', ftString, 60);
  bufData.FieldDefs.Add('Jabatan', ftString, 40);
  bufData.FieldDefs.Add('Departemen', ftString, 40);
  bufData.FieldDefs.Add('Kantor', ftString, 40);
  bufData.FieldDefs.Add('Tanggal', ftDate);
  bufData.FieldDefs.Add('Scan1', ftString, 8);
  bufData.FieldDefs.Add('Scan2', ftString, 8);
  bufData.FieldDefs.Add('Scan3', ftString, 8);
  bufData.FieldDefs.Add('JamKerja', ftString, 8);
  bufData.FieldDefs.Add('Menit', ftInteger);
  bufData.CreateDataset;
  bufData.Open;
  FormatGrids;
end;

procedure TfrmMain.PrepareRekap;
begin
  bufRekap.Close;
  bufRekap.FieldDefs.Clear;
  bufRekap.FieldDefs.Add('PIN', ftString, 20);
  bufRekap.FieldDefs.Add('Nama', ftString, 60);
  bufRekap.FieldDefs.Add('HariHadir', ftInteger);
  bufRekap.FieldDefs.Add('HariTakLengkap', ftInteger);
  bufRekap.FieldDefs.Add('TotalJam', ftString, 10);
  bufRekap.FieldDefs.Add('RataJam', ftString, 10);
  bufRekap.CreateDataset;
  bufRekap.Open;
  FormatGrids;
end;

procedure TfrmMain.FormatGrids;
begin
  if bufData.Active then
  begin
    TDateField(bufData.FieldByName('Tanggal')).DisplayFormat := 'dd-mm-yyyy';
    bufData.FieldByName('Scan1').DisplayLabel := 'Scan 1';
    bufData.FieldByName('Scan2').DisplayLabel := 'Scan 2';
    bufData.FieldByName('Scan3').DisplayLabel := 'Scan 3';
    bufData.FieldByName('JamKerja').DisplayLabel := 'Jam Kerja';
    bufData.FieldByName('Menit').Visible := False;
  end;
  if bufRekap.Active then
  begin
    bufRekap.FieldByName('HariHadir').DisplayLabel := 'Hari Hadir';
    bufRekap.FieldByName('HariTakLengkap').DisplayLabel := 'Scan Tak Lengkap';
    bufRekap.FieldByName('TotalJam').DisplayLabel := 'Total Jam Kerja';
    bufRekap.FieldByName('RataJam').DisplayLabel := 'Rata2 Jam/Hari';
  end;
end;

procedure TfrmMain.UpdateStatus;
var
  s: string;
begin
  if FFileName = '' then
    s := 'Belum ada data. Klik "Buka CSV..." untuk import.'
  else
    s := Format('File: %s   |   Baris data: %d   |   Pegawai: %d',
      [FFileName, bufData.RecordCount, bufRekap.RecordCount]);
  sb.SimpleText := s;
end;

procedure TfrmMain.btnOpenClick(Sender: TObject);
begin
  if dlgOpen.Execute then
  begin
    try
      LoadCSV(dlgOpen.FileName);
    except
      on E: Exception do
        MessageDlg('Gagal import', E.Message, mtError, [mbOK], 0);
    end;
  end;
end;

procedure TfrmMain.LoadCSV(const AFileName: string);
var
  sl, hdr, row: TStringList;
  i, hdrLine, n, skipped: Integer;
  delim: Char;
  cPIN, cNIP, cNama, cJab, cDep, cKan, cTgl, cS1, cS2, cS3: Integer;
  tgl: TDateTime;
  scans: array[1..3] of string;
  k, mn, mnMin, mnMax, cnt: Integer;
  line: string;
begin
  sl := TStringList.Create;
  hdr := TStringList.Create;
  row := TStringList.Create;
  Screen.Cursor := crHourGlass;
  try
    sl.LoadFromFile(AFileName);
    if sl.Count = 0 then
      raise Exception.Create('File kosong.');

    { buang BOM UTF-8 bila ada }
    if Copy(sl[0], 1, 3) = #$EF#$BB#$BF then
      sl[0] := Copy(sl[0], 4, MaxInt);

    { cari baris header (yang memuat "PIN" dan "Tanggal") }
    hdrLine := -1;
    for i := 0 to sl.Count - 1 do
    begin
      line := LowerCase(sl[i]);
      if (Pos('pin', line) > 0) and (Pos('tanggal', line) > 0) then
      begin
        hdrLine := i;
        Break;
      end;
    end;
    if hdrLine < 0 then
      raise Exception.Create('Baris header tidak ditemukan.' + LineEnding +
        'Kolom yang diharapkan: PIN, NIP, Nama, Jabatan, Departemen, Kantor, ' +
        'Tanggal, Scan 1, Scan 2, Scan 3');

    delim := DetectDelimiter(sl[hdrLine]);
    SplitCSV(sl[hdrLine], delim, hdr);

    cPIN := FindCol(hdr, 'PIN');
    cNIP := FindCol(hdr, 'NIP');
    cNama := FindCol(hdr, 'Nama');
    cJab := FindCol(hdr, 'Jabatan');
    cDep := FindCol(hdr, 'Departemen');
    cKan := FindCol(hdr, 'Kantor');
    cTgl := FindCol(hdr, 'Tanggal');
    cS1 := FindCol(hdr, 'Scan 1');
    cS2 := FindCol(hdr, 'Scan 2');
    cS3 := FindCol(hdr, 'Scan 3');

    if (cPIN < 0) or (cTgl < 0) or (cS1 < 0) then
      raise Exception.Create('Kolom PIN / Tanggal / Scan 1 tidak ditemukan di header.');

    { kosongkan data lama }
    edtCari.OnChange := nil;
    edtCari.Text := '';
    edtCari.OnChange := @edtCariChange;
    PrepareDataset;

    n := 0;
    skipped := 0;
    bufData.DisableControls;
    try
      for i := hdrLine + 1 to sl.Count - 1 do
      begin
        if Trim(sl[i]) = '' then Continue;
        SplitCSV(sl[i], delim, row);

        if not ParseTanggal(GetCol(row, cTgl), tgl) then
        begin
          Inc(skipped);
          Continue;
        end;

        scans[1] := GetCol(row, cS1);
        scans[2] := GetCol(row, cS2);
        scans[3] := GetCol(row, cS3);

        { jam kerja = scan terakhir - scan pertama }
        mnMin := MaxInt;
        mnMax := -1;
        cnt := 0;
        for k := 1 to 3 do
        begin
          mn := ParseJamMenit(scans[k]);
          if mn >= 0 then
          begin
            Inc(cnt);
            if mn < mnMin then mnMin := mn;
            if mn > mnMax then mnMax := mn;
          end;
        end;

        bufData.Append;
        bufData.FieldByName('PIN').AsString := GetCol(row, cPIN);
        bufData.FieldByName('NIP').AsString := GetCol(row, cNIP);
        bufData.FieldByName('Nama').AsString := GetCol(row, cNama);
        bufData.FieldByName('Jabatan').AsString := GetCol(row, cJab);
        bufData.FieldByName('Departemen').AsString := GetCol(row, cDep);
        bufData.FieldByName('Kantor').AsString := GetCol(row, cKan);
        bufData.FieldByName('Tanggal').AsDateTime := tgl;
        bufData.FieldByName('Scan1').AsString := scans[1];
        bufData.FieldByName('Scan2').AsString := scans[2];
        bufData.FieldByName('Scan3').AsString := scans[3];
        if cnt >= 2 then
        begin
          bufData.FieldByName('Menit').AsInteger := mnMax - mnMin;
          bufData.FieldByName('JamKerja').AsString := MenitToStr(mnMax - mnMin);
        end
        else
        begin
          bufData.FieldByName('Menit').AsInteger := 0;
          bufData.FieldByName('JamKerja').AsString := '';
        end;
        bufData.Post;
        Inc(n);
      end;
      bufData.First;
    finally
      bufData.EnableControls;
    end;

    FFileName := AFileName;
    BuildRekap;
    gridData.AutoAdjustColumns;
    gridRekap.AutoAdjustColumns;
    UpdateStatus;

    if skipped > 0 then
      ShowMessage(Format('%d baris berhasil diimport, %d baris dilewati ' +
        '(tanggal tidak valid).', [n, skipped]));
  finally
    Screen.Cursor := crDefault;
    row.Free;
    hdr.Free;
    sl.Free;
  end;
end;

procedure TfrmMain.BuildRekap;
var
  list: TStringList;
  r: TRekap;
  idx, i, mn: Integer;
  key, pin: string;
  bm: TBookmark;
begin
  PrepareRekap;
  list := TStringList.Create;
  try
    list.Sorted := True;
    list.Duplicates := dupIgnore;

    bufData.DisableControls;
    bm := bufData.Bookmark;
    try
      bufData.First;
      while not bufData.EOF do
      begin
        pin := bufData.FieldByName('PIN').AsString;
        key := Format('%.10d|%s', [StrToIntDef(pin, 0), pin]);
        idx := list.IndexOf(key);
        if idx < 0 then
        begin
          r := TRekap.Create;
          r.PIN := pin;
          r.Nama := bufData.FieldByName('Nama').AsString;
          idx := list.AddObject(key, r);
        end;
        r := TRekap(list.Objects[idx]);
        Inc(r.Hari);
        mn := bufData.FieldByName('Menit').AsInteger;
        if bufData.FieldByName('JamKerja').AsString <> '' then
        begin
          Inc(r.HariLengkap);
          Inc(r.Menit, mn);
        end;
        bufData.Next;
      end;
      bufData.First;
    finally
      bufData.EnableControls;
    end;

    bufRekap.DisableControls;
    try
      for i := 0 to list.Count - 1 do
      begin
        r := TRekap(list.Objects[i]);
        bufRekap.Append;
        bufRekap.FieldByName('PIN').AsString := r.PIN;
        bufRekap.FieldByName('Nama').AsString := r.Nama;
        bufRekap.FieldByName('HariHadir').AsInteger := r.Hari;
        bufRekap.FieldByName('HariTakLengkap').AsInteger := r.Hari - r.HariLengkap;
        bufRekap.FieldByName('TotalJam').AsString := MenitToStr(r.Menit);
        if r.HariLengkap > 0 then
          bufRekap.FieldByName('RataJam').AsString :=
            MenitToStr(r.Menit div r.HariLengkap)
        else
          bufRekap.FieldByName('RataJam').AsString := '';
        bufRekap.Post;
      end;
      bufRekap.First;
    finally
      bufRekap.EnableControls;
    end;
  finally
    for i := 0 to list.Count - 1 do
      list.Objects[i].Free;
    list.Free;
  end;
end;

procedure TfrmMain.bufDataFilterRecord(DataSet: TDataSet; var Accept: Boolean);
var
  q: string;
begin
  q := UpperCase(Trim(edtCari.Text));
  if q = '' then
    Accept := True
  else
    Accept := (Pos(q, UpperCase(DataSet.FieldByName('Nama').AsString)) > 0) or
              (Pos(q, UpperCase(DataSet.FieldByName('PIN').AsString)) > 0);
end;

procedure TfrmMain.edtCariChange(Sender: TObject);
begin
  if not bufData.Active then Exit;
  bufData.Filtered := False;
  bufData.Filtered := Trim(edtCari.Text) <> '';
end;

procedure TfrmMain.btnExportClick(Sender: TObject);
var
  sl: TStringList;
  i: Integer;
  s: string;
  fld: TField;
const
  Delim = ';'; // pemisah ';' agar langsung rapi di Excel regional Indonesia
begin
  if (not bufData.Active) or (bufData.IsEmpty) then
  begin
    ShowMessage('Tidak ada data untuk diexport.');
    Exit;
  end;
  if not dlgSave.Execute then Exit;

  sl := TStringList.Create;
  bufData.DisableControls;
  try
    sl.Add('PIN;NIP;Nama;Jabatan;Departemen;Kantor;Tanggal;Scan 1;Scan 2;Scan 3;Jam Kerja');
    bufData.First;
    while not bufData.EOF do
    begin
      s := '';
      for i := 0 to 10 do
      begin
        fld := bufData.Fields[i];
        if i > 0 then s := s + Delim;
        if fld.FieldName = 'Tanggal' then
          s := s + FormatDateTime('dd-mm-yyyy', fld.AsDateTime)
        else
          s := s + CsvQuote(fld.AsString, Delim);
      end;
      sl.Add(s);
      bufData.Next;
    end;
    bufData.First;
    sl.SaveToFile(dlgSave.FileName);
    ShowMessage(Format('%d baris berhasil diexport.', [sl.Count - 1]));
  finally
    bufData.EnableControls;
    sl.Free;
  end;
end;

end.
