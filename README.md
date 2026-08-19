# Vale-Sêmen Municipal

Aplicativo em Google Apps Script para calcular, emitir, registrar e reimprimir vales de sêmen destinados a produtores rurais.

## Recursos

- cálculo por faixas de produção anual;
- limite configurável de doses;
- numeração anual automática com bloqueio contra concorrência;
- prevenção de emissão duplicada por CPF e ano;
- geração de Google Docs e PDF;
- senha de emissão armazenada nas propriedades do script;
- registro em Google Sheets.

## Instalação

1. Crie uma planilha com as abas `Regras` e `Registros`.
2. Crie um modelo no Google Docs e uma pasta para os documentos gerados.
3. Crie um projeto no Google Apps Script e adicione `Code.gs` e `Formulario.html`.
4. Troque os valores de exemplo em `CONFIG` pelos IDs dos seus próprios recursos.
5. Defina `SENHA_EMISSAO` em **Configurações do projeto > Propriedades do script**.
6. Execute `prepararSistema()` uma vez e autorize o projeto.
7. Implante como aplicativo da Web.

## Privacidade

Este repositório não contém planilhas, documentos, CPFs, senhas ou identificadores do ambiente em produção. Cada instalação deve usar recursos próprios e observar a LGPD.

## Licença

MIT. Consulte [LICENSE](LICENSE).

