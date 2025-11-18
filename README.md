# Validador de PDF

Este pacote contém um script em Python que valida PDFs/imagens segundo o modelo de
"DECLARAÇÃO DE RECEBIMENTO DE MATERIAL DE LIMPEZA" e grava o resultado em uma planilha Excel.

Arquivos incluídos (em `deploy_files.zip`):

- `validate_pdf_standard.py` — script principal (Python)
- `requirements.txt` — dependências Python (instalar com `pip`)
- `README.md` — este guia passo a passo
- `learned_patterns.json` — dados de aprendizado (modelos já reconhecidos)
- `por.traineddata` — dados de idioma Português para o Tesseract (caso sua instalação não o contenha)
- `run_validate.bat` — atalho para executar o script com dois cliques

---

Guia rápido — passo a passo para quem nunca programou
=======================================================

Siga cada passo com atenção. Se algo não funcionar, copie a mensagem do erro e peça ajuda.

1) Extraia os arquivos

   - Clique com o botão direito em `deploy_files.zip` → `Extrair Tudo...` → escolha uma pasta (ex.: `C:\Users\%USERNAME%\Documents\Leitor de PDF`) → `Extrair`.
2) Instale o Python (se ainda não tiver)

   - Abra [https://www.python.org/downloads/](https://www.python.org/downloads/)
   - Baixe o instalador recomendado (Python 3.11+).
   - Execute o instalador e marque **Add Python to PATH** antes de clicar em **Install Now**.
3) Abra o Prompt de Comando (cmd)

   # Validador de PDF — Guia rápido (usuário leigo)

   Este pacote contém um script Python que valida PDFs/imagens segundo o modelo
   "DECLARAÇÃO DE RECEBIMENTO DE MATERIAL DE LIMPEZA" e grava os resultados numa planilha Excel.

   Arquivos incluídos no `deploy_files.zip`:

   - `validate_pdf_standard.py` — script principal
   - `requirements.txt` — lista de dependências Python
   - `README.md` — este guia
   - `learned_patterns.json` — padrões aprendidos
   - `por.traineddata` — dado de idioma Português para Tesseract
   - `run_validate.bat` — atalho para executar tudo por duplo clique

   Importante sobre o `run_validate.bat`:

   - O `run_validate.bat` verifica cada pacote listado em `requirements.txt`. Se algum pacote estiver faltando, ele instala apenas o pacote necessário usando `pip --user`. Se todos já estiverem instalados, nada é instalado. O `.bat` NÃO instala o Tesseract (é um binário nativo); instruções para isso estão abaixo.

   Passo a passo (curto)
   ---------------------

   1) Extraia `deploy_files.zip` numa pasta (ex.: `C:\Users\%USERNAME%\Documents\Leitor de PDF`).

   2) Instale o Python se necessário: baixe em https://www.python.org/downloads/ e marque **Add Python to PATH** durante a instalação.

   3) (Opcional) Abra um `cmd` e vá para a pasta do projeto:

   ```cmd
   cd "C:\Users\%USERNAME%\Documents\Leitor de PDF"
   ```

   4) Duplo clique em `run_validate.bat` para:
   - verificar/instalar só os pacotes Python que faltarem; e
   - executar `validate_pdf_standard.py` ao final.

   5) Se preferir rodar manualmente (linha de comando):

   ```cmd
   python validate_pdf_standard.py "C:\caminho\para\sua_planilha.xlsx" --tesseract-cmd "C:\Program Files\Tesseract-OCR\tesseract.exe"
   ```

   Verificando o Tesseract (OCR)
   ----------------------------

   - Instale Tesseract separadamente (ex.: UB-Mannheim build) — o `.bat` não faz essa instalação.
   - Teste no `cmd`:

   ```cmd
   tesseract --version
   tesseract --list-langs
   ```

   Se `por` não estiver disponível, o `por.traineddata` incluído no ZIP será usado automaticamente pelo script (ele define `TESSDATA_PREFIX` quando encontra o arquivo no diretório do projeto).

   Problemas comuns
   ----------------

   - `tesseract: command not found` → instale o Tesseract ou use `--tesseract-cmd` para apontar o executável.
   - `Arquivo não encontrado` (caminho UNC) → copie PDFs para uma pasta local antes de rodar.

   Se quiser, posso também substituir `requirements.txt` por uma versão limpa (eu já tenho uma `requirements_clean.txt` no pacote). Neste momento vou substituir o `requirements.txt` por uma versão limpa.

"C:\...\python.exe" validate_pdf_standard.py "C:\Users\victor.vasconcelos\Documents\Leitor de PDF\material de limpeza.xlsx"
```

## 3) Tesseract (OCR) — instalação e apontamento

Observação: o Tesseract é um binário nativo (não instalado via pip). O pacote `pytesseract` no `requirements.txt` permite controlar o binário, mas o executável precisa estar instalado.

Opções de instalação no Windows (passo a passo):

- a) Via Scoop (sem privilégios de administrador):

  Abra PowerShell como usuário normal e rode:

  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
  iwr -useb get.scoop.sh | iex
  scoop install tesseract
  ```

  Após isso, verifique no cmd:

  ```cmd
  tesseract --version
  ```
- b) Via Chocolatey (se você tiver Chocolatey):

  ```powershell
  choco install tesseract
  ```
- c) Via Conda (se preferir instalar dentro do ambiente conda):

  ```cmd
  conda install -c conda-forge tesseract pytesseract pillow
  ```
- d) Build portátil / ZIP (baixar manualmente):

  - Baixe uma build para Windows (ex.: repositório oficial: https://github.com/tesseract-ocr/tesseract; builds Windows comuns: https://github.com/UB-Mannheim/tesseract/wiki )
  - Extraia em `C:\Users\<seu_user>\tools\tesseract\` e garanta a pasta `tessdata` com `por.traineddata` presente.

Depois de instalar, você pode apontar explicitamente o executável quando chamar o script usando a flag `--tesseract-cmd`:

```cmd
"C:\...\python.exe" validate_pdf_standard.py "C:\path\to\sheet.xlsx" --tesseract-cmd "C:\Users\me\tools\tesseract\tesseract.exe"
```

O script também tenta detectar `tesseract` automaticamente em locais comuns (Scoop, conda, PATH). Se não encontrar, e se estiver rodando em um terminal interativo, ele perguntará pelo caminho do `tesseract.exe`.

## 4) Exemplo completo de execução (Windows, cmd.exe)

- Rodar interativo (script perguntará planilha e tesseract se necessário):

```cmd
"C:\...\python.exe" validate_pdf_standard.py
```

- Rodar não interativo (passando planilha e tesseract):

```cmd
"C:\...\python.exe" validate_pdf_standard.py "C:\Users\victor.vasconcelos\Documents\Leitor de PDF\material de limpeza.xlsx" --tesseract-cmd "C:\Users\victor.vasconcelos\scoop\shims\tesseract.exe"
```

- Se quiser pular cabeçalho diferente (0 linhas de cabeçalho):

```cmd
... validate_pdf_standard.py "...\material de limpeza.xlsx" --header-rows 0
```

## 5) Saída e colunas da planilha

- O script lê caminhos na **coluna G** (coluna 7) e grava o resultado em **coluna M** (coluna 13).
- O resultado na coluna M será **"Documento Aprovado"** ou **"Documento Reprovado"**. A coluna N (14) contém notas/audit trail (por exemplo: motivo da reprovação ou se foi auto-aprovado pelo aprendizado).

## 6) Boas práticas para performance (sem alterar lógica)

- Se os arquivos estão em um servidor de rede (UNC) e o processamento for grande, copie os PDFs localmente antes de rodar para reduzir latência (ex.: `C:\data\dec_kli`). Use `robocopy` para grandes cópias.
- Evite abrir/salvar o Excel a cada arquivo: prefira rodar o script como está (ele já aplica saves periódicos), ou use modos em lote/parallel conforme necessário (avançado).

## 7) Dicas de depuração

- Se o OCR falhar por ausência de idioma em português, coloque `por.traineddata` em `tessdata` do executável Tesseract ou na pasta do script; o script também tentará definir `TESSDATA_PREFIX` automaticamente quando encontrar esse arquivo.
- Se o script não encontrar `tesseract` automaticamente, rode com `--tesseract-cmd` apontando o executável.

Arquivo: `validate_pdf_standard.py` — mantém toda a lógica de validação, OCR forçado e aprendizado (`learned_patterns.json`). Este README apenas documenta como preparar e executar o script de forma simples e reproduzível.
