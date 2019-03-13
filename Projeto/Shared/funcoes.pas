unit Funcoes;

// BIBLIOTECA DE FUNÇÕES
// DESENVOLVIDA POR VALTER PATRICK

interface

uses
  System.SysUtils, System.Variants, System.Classes, System.StrUtils, System.Math, System.DateUtils,
  Windows, Forms;

function FillChar(Str: string; Len: Integer; Ch: Char = ' '; Right: Boolean = False): String;
function SomenteNumeros(Str: String): String;
function StrInSet(StrChar: String; StrSet: array of string): Boolean;
function StringInSet(Str: String; StrSet: array of string): Boolean;
function IsNumber(Str: String): Boolean;
function TemCaracteresEspeciais(Str: String): Boolean;
function VersaoExe: String;

implementation

function FillChar(Str: string; Len: Integer; Ch: Char = ' '; Right: Boolean = False): String;
var
  I: Integer;
begin
  Result := Str.Trim;
  if Right then // Direito
  begin
    for I := Length(Str) to Len - 1 do
      Result := Result + Ch;
  end
  else
  begin // Esquerdo
    for I := Length(Str) to Len - 1 do
      Result := Ch + Result;
  end;
end;

function SomenteNumeros(Str: String): String;
var
  I: Integer;
begin
  Result := '';
  if Trim(Str) = '' then
    Exit;
  for I := 1 to Length(Str) do
  begin
    if CharInSet(Str[I], ['0' .. '9']) then
      Result := Result + Str[I];
  end;
end;

function StrInSet(StrChar: String; StrSet: array of string): Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := Low(StrSet) to High(StrSet) do
  begin
    if StrSet[I] = StrChar then
    begin
      Result := True;
      Break;
    end;
  end;
end;

function StringInSet(Str: String; StrSet: array of string): Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := 1 to Length(Str) do
  begin
    if StrInSet(Str[I], StrSet) then
    begin
      Result := True;
      Break;
    end;
  end;
end;

function IsNumber(Str: String): Boolean;
var
  I: Integer;
begin
  Result := Length(Str) > 0;
  for I := 1 to Length(Str) do
  begin
    if not(CharInSet(Str[I], ['0' .. '9'])) then
    begin
      Result := False;
      Break;
    end;
  end;
end;

function TemCaracteresEspeciais(Str: String): Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := 1 to Length(Str) do
  begin
    if not CharInSet(Str[I], ['a' .. 'z', 'A' .. 'Z', '0' .. '9', ' ']) then
    begin
      Result := True;
      Break;
    end;
  end;
end;

function VersaoExe: String;
type
  PFFI = ^vs_FixedFileInfo;
var
  F: PFFI;
  Handle: Dword;
  Len: Longint;
  Data: Pchar;
  Buffer: Pointer;
  Tamanho: Dword;
  Parquivo: Pchar;
  Arquivo: String;
begin
  Arquivo := Application.ExeName;
  Parquivo := StrAlloc(Length(Arquivo) + 1);
  StrPcopy(Parquivo, Arquivo);
  Len := GetFileVersionInfoSize(Parquivo, Handle);
  Result := '';
  if Len > 0 then
  begin
    Data := StrAlloc(Len + 1);
    if GetFileVersionInfo(Parquivo, Handle, Len, Data) then
    begin
      VerQueryValue(Data, '', Buffer, Tamanho);
      F := PFFI(Buffer);
      Result := Format('%d.%d.%d.%d', [HiWord(F^.dwFileVersionMs), LoWord(F^.dwFileVersionMs), HiWord(F^.dwFileVersionLs), LoWord(F^.dwFileVersionLs)]);
    end;
    StrDispose(Data);
  end;
  StrDispose(Parquivo);
  // Referência:
  // http://www.planetadelphi.com.br/dica/6132/funcao-que-retorna-a-versao-do-proprio-programa
end;

end.
