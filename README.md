# 📋 Preenchedor Automático de Formulários (Microsoft Forms)

Este projeto é um script em Python que automatiza o preenchimento de formulários hospedados no **Microsoft Forms**, com base em dados extraídos de uma planilha Excel.

## 🚀 Funcionalidades

- Preenche campos de texto, datas e menus suspensos (dropdowns);
- Verifica se a opção existe antes de selecionar (dropdowns);
- Lê uma planilha Excel com múltiplas linhas, preenchendo um formulário para cada uma;
- Gera capturas de tela automáticas em caso de erro;
- Usa o Selenium para navegação e preenchimento automático.

## 🔧 Requisitos

- Python 3.9 ou superior
- Google Chrome (instalado)
- ChromeDriver compatível com sua versão do Chrome
- Biblioteca Selenium (`pip install selenium`)
- Biblioteca Pandas (`pip install pandas`)
- Biblioteca OpenPyXL (`pip install openpyxl`)

...

## ▶️ Como usar

1. Clone este repositório ou baixe os arquivos.
2. Certifique-se de ter o Chrome e o ChromeDriver instalados.
3. Preencha a planilha `BaseRH.xlsx` com os dados necessários.
4. Execute o script:

```bash
python fill_ms_forms_dynamic.py
```

O script abrirá o navegador, acessará o formulário do Microsoft Forms e preencherá os campos de acordo com cada linha da planilha.

## 🧠 Observações Técnicas

- O formulário deve estar público (sem necessidade de login);
- Os campos devem manter a mesma ordem e estrutura da planilha;
- O script trabalha com dropdowns reais (não apenas `select` HTML, mas elementos clicáveis do Forms).

## 📌 Observações

> Este projeto foi desenvolvido com fins empresariais no contexto do time de RH da **Âncora Consórcios**.  
> Seu objetivo principal é automatizar o preenchimento de formulários a partir de planilhas internas.  
> Pode ser adaptado para outros contextos conforme necessário.

## 🧑‍💻 Autor
Gabriel Piccirillo
