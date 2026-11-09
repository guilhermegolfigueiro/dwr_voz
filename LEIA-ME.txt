# DWR Voice — ditado no iPhone -> DWR completo no Word

## O que o sistema faz
- O iPhone salva textos ditados no iCloud Drive, em DWR_Entrada/fila.
- ACT_*.txt = linhas da tabela LIST OF ACTIVITIES PERFORMED.
- REQ_*.txt = linhas da tabela ACTIONS / REQUESTS.
- O programa no Windows monitora a fila e gera, a cada ditado, o DWR completo do dia:
  cabecalho e assinaturas copiados do documento base + numero + data + as duas tabelas.
- O DWR e salvo em OneDrive/DWR/relatorios como DWR_dd-mm-aaaa.docx.
- Ditados processados vao para "feito"; os que falham vao para "erros".

## Arquivos desta pasta
- dwr_voice.py   programa principal (deixe aberto no Windows)
- gerar_dwr.py   monta o Word a partir do base (precisa estar na mesma pasta)
- config.json    caminhos (so preencha se a deteccao automatica falhar)
- iniciar_dwr.bat  atalho para iniciar
- exemplo_base/base.docx  modelo de documento base (so um exemplo)
- extras/processar_fila.py  versao alternativa para a-Shell (nao e necessaria)

## Instalacao no Windows
1. Instale o Python 3 (python.org ou Microsoft Store). Nao ha nenhum pacote para instalar.
2. Instale o iCloud para Windows, entre com a mesma Conta Apple do iPhone e ative iCloud Drive.
3. Confirme que existe a pasta OneDrive/DWR (o programa cria "relatorios" dentro dela).
4. Extraia esta pasta, por exemplo em C:\dwr_voz.
5. Execute iniciar_dwr.bat e deixe a janela aberta.
   Na primeira vez, o programa cria em iCloud Drive/DWR_Entrada as pastas fila, feito, erros e exemplo.
6. Se a deteccao automatica falhar, edite config.json:
   - "icloud_drive": caminho da pasta iCloud Drive
   - "onedrive_dwr": caminho completo de OneDrive\DWR
   - "formato_data" (opcional): formato da data escrita dentro do DWR. Padrao "%m/%d/%Y".
     Para dia/mes/ano use "%d/%m/%Y".

## Documento base (a cada novo servico)
1. Monte o DWR do servico com cabecalho, tabelas e assinaturas preenchidos.
2. Salve como base.docx em iCloud Drive/DWR_Entrada/exemplo
   (se for o unico .docx da pasta, qualquer nome serve).
3. O numero do base e o do primeiro DWR; a cada novo dia sobe 1 e os registros recomecam.
4. Para iniciar outro servico: troque o base.docx e apague o arquivo estado.json
   (fica na pasta do programa), para o numero recomecar.

O base precisa ter: 6 tabelas (a 1a com o numero e a data dd/dd/aaaa), os titulos
"LIST OF ACTIVITIES PERFORMED" e "ACTIONS / REQUESTS" e duas tabelas com cabecalho Item / Description.
Linhas de nota com "*" (ex.: *Brazilian time) nao sao mantidas no DWR gerado.

## Atalhos no iPhone
Crie dois atalhos: Ditado de Texto (Portugues) -> Texto -> Salvar Arquivo em iCloud Drive/DWR_Entrada/fila.
- "DWR - Atividades": nome ACT_[yyyy-MM-dd-HHmmss].txt
- "DWR - Acoes": nome REQ_[yyyy-MM-dd-HHmmss].txt
Toque Atras (Ajustes > Acessibilidade > Toque > Tocar Atras): duplo = Atividades, triplo = Acoes.

## Observacoes
- Cada ditado regenera o Word inteiro: edicoes manuais no DWR do dia sao perdidas no proximo ditado.
  Edite so depois de terminar o dia.
- Se o DWR do dia estiver aberto no Word ou nao houver base, o programa avisa e deixa os ditados na fila.
- Mantenha so UMA janela do programa aberta (duas instancias duplicam registros).
- O computador precisa estar ligado, online e com o programa em execucao.
- A data do relatorio e a data local do computador (America/Sao_Paulo).
