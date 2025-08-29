unit uMensagens;

interface

type
  TMsg = class
    class procedure Alerta(texto: string; titulo: string = 'Atenção');
    class procedure Erro(texto: string; titulo: string = 'Erro');
    class procedure Informacao(texto: string; titulo: string = 'Informação');
    class function Confirmacao(texto: string; titulo: string = 'Confirmação'; naoComoPadrao: Boolean = false): Boolean;
    class procedure Beep();
  end;

implementation

uses Forms, Windows;

{ TMsg }

class procedure TMsg.Alerta(texto: string; titulo: string = 'Atenção');
begin

  Application.MessageBox(PChar(texto), PChar(titulo), MB_OK + MB_ICONWARNING);
end;

class procedure TMsg.Beep;
begin

  Windows.Beep(4000, 500);

end;

class function TMsg.Confirmacao(texto: string; titulo: string = 'Confirmação'; naoComoPadrao: Boolean = false): Boolean;
begin

  if (naoComoPadrao) then
  begin

    Result := Application.MessageBox(PChar(texto), PChar(titulo), MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2) = IDYES;
  end
  else
  begin

    Result := Application.MessageBox(PChar(texto), PChar(titulo), MB_YESNO + MB_ICONQUESTION) = IDYES;
  end;

end;

class procedure TMsg.Erro(texto: string; titulo: string = 'Erro');
begin

  Application.MessageBox(PChar(texto), PChar(titulo), MB_OK + MB_ICONERROR);
end;

class procedure TMsg.Informacao(texto: string; titulo: string = 'Informação');
begin

  Application.MessageBox(PChar(texto), PChar(titulo), MB_OK + MB_ICONINFORMATION);
end;

end.
