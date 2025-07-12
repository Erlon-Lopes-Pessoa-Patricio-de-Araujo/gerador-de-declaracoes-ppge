# 🧾 Gerador de Declarações - PPGE

![Python](https://img.shields.io/badge/Python-3.x-blue?logo=python)
![License](https://img.shields.io/badge/License-MIT-green)

### ℹ️ Contexto

Eu tinha um problema recorrente no PPGE: recebia uma **relação de alunos** com os **nomes e matrículas**, mas os **CPFs precisavam ser coletados manualmente** depois, um por um. Para **adiantar esse processo** e evitar retrabalho na hora de gerar declarações personalizadas, criei esse sistema simples que automatiza tudo.

Ele me permite:

* Selecionar um modelo de declaração `.docx`
* Ler automaticamente os nomes e matrículas da relação
* Coletar os CPFs direto pelo terminal
* E gerar os arquivos prontos em segundos, com tudo nomeado e organizado

---


Este é um script em Python desenvolvido para automatizar a geração de declarações personalizadas para os(as) alunos(as) de uma turma, utilizando um modelo `.docx` e uma relação de alunos extraída de outro documento Word.

Ideal para situações em que é necessário gerar rapidamente várias declarações com informações como **nome**, **CPF** e **matrícula** de forma automatizada.

---

## 🚀 Funcionalidades

- Interface gráfica para seleção de arquivos via `tkinter`
- Leitura de relação de alunos em tabelas do Word
- Substituição de texto no modelo mantendo a formatação
- Validação e formatação de CPF
- Salvamento automático com nome de arquivo personalizado
- Criação automática da pasta de saída
- Abertura automática da pasta ao final do processo

---

## 📁 Estrutura esperada

- **Modelo de declaração (`.docx`)** deve conter os seguintes marcadores:
  - `xxxxxxxxxxxxxxxxxxxxx` → será substituído pelo **nome**
  - `xxxxxxxxxxxxxx` → será substituído pelo **CPF**
  - `xxxxx` → será substituído pela **matrícula**

- **Relação de alunos (`.docx`)** deve conter uma tabela com:
  - Matrícula (primeira coluna)
  - Nome do aluno (segunda coluna)

---

## ▶️ Como usar

1. Execute o script:

```bash
python autoação_ppge.py
````

2. Selecione o modelo da declaração (`.docx`)
3. Selecione a relação de alunos com tabela (`.docx`)
4. Digite os CPFs conforme solicitado no terminal
5. Aguarde a geração automática dos arquivos

---

## 📦 Requisitos

* Python 3.x
* Pacotes:

  * `python-docx`
  * `tkinter` (já incluso no Python em muitos sistemas)
  * `unicodedata`, `re`, `os` (bibliotecas padrão)

Instale o `python-docx` caso não tenha:

```bash
pip install python-docx
```

---

## 🧪 Exemplo de uso

Você pode testar com os arquivos da pasta `exemplos/`:

* [`modelo-exemplo.docx`]([exemplos/modelo-exemplo.docx](https://github.com/Erlon-Lopes-Pessoa-Patricio-de-Araujo/gerador-de-declaracoes-ppge/compare/main...Exemplos#diff-cff880e0482624352a657e50cf1cf2d370100a5ddfbc79b5e16f9641da6de561))
* [`relacao-exemplo.docx`]([exemplos/relacao-exemplo.docx](https://github.com/Erlon-Lopes-Pessoa-Patricio-de-Araujo/gerador-de-declaracoes-ppge/compare/main...Exemplos#diff-2cac77c609bad191929f4f78372e0db4ae0e6372fda5dc882b9903ef6b7f7c45))

---

## 👤 Autor

Desenvolvido por **Erlon Lopes** com foco em produtividade e praticidade para uso acadêmico no PPGE.

---

## 📄 Licença

Este projeto está sob a licença MIT. Sinta-se livre para usar, modificar e contribuir.

```


