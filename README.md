# POP Zero

Aplicação estática para extrair o conteúdo de POPs antigos em `.docx`, organizar as seções conforme a Norma Zero e gerar um novo Word completo no padrão institucional.

## Recursos

- Processamento inteiramente no navegador.
- Identificação automática das oito seções da Norma Zero.
- Revisão e edição antes da exportação.
- Preservação de tabelas como elementos editáveis e recuperação de tabelas antigas montadas com tabulações.
- Formulário para título, processo, código, versão, elaboração, aprovação, datas e vigência.
- Exportação em `.docx` com Arial 12, cabeçalho institucional, sumário automático, histórico das revisões, rodapé e paginação.
- Compatível com GitHub Pages.

## Publicar no GitHub Pages

1. Crie um repositório no GitHub.
2. Envie `index.html`, `styles.css` e `app.js` para a raiz do repositório.
3. Abra **Settings > Pages**.
4. Em **Build and deployment**, escolha **Deploy from a branch**.
5. Selecione a branch `main`, a pasta `/root` e clique em **Save**.

O endereço será disponibilizado pelo GitHub após a publicação.

## Executar localmente

Abra `index.html` no navegador. Também é possível iniciar um servidor simples:

```bash
python -m http.server 8000
```

Depois acesse `http://localhost:8000`.

## Privacidade

O arquivo DOCX é processado localmente. Nenhum conteúdo do documento é enviado a um servidor pela aplicação.

## Limitações

- O reconhecimento depende dos títulos das seções. Variações muito diferentes podem exigir ajuste manual.
- Elementos complexos, como imagens e fluxogramas, não são importados automaticamente.
- A tela de revisão permite corrigir ou complementar qualquer seção antes da geração.
- O Microsoft Word poderá solicitar a atualização do sumário ao abrir o arquivo; confirme a atualização para recalcular os números das páginas.

## Bibliotecas utilizadas

- JSZip 3.10.1 para leitura do pacote DOCX.
- docx 8.5.0 para geração do arquivo Word.

As bibliotecas estão incluídas na pasta `vendor`. Depois de publicada, a aplicação funciona sem serviços externos e sem enviar o conteúdo dos documentos para terceiros.
