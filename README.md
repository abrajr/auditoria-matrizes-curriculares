# Auditoria Regulatória de Matrizes Curriculares

Ferramenta de **validação de carga horária de matrizes curriculares** segundo as
regras do MEC e as Diretrizes Curriculares Nacionais (DCNs), usando o solver
**CP-SAT do Google OR-Tools**, com interface gráfica em Tkinter e relatório
executivo em Excel.

> ⚠️ **ATENÇÃO — DADOS FICTÍCIOS**
> A base de exemplo que acompanha a ferramenta (`Dados_Exportacao_CSV.csv`) é
> **sintética**, criada apenas para demonstrar o funcionamento do código. Os
> cursos, matrizes e cargas horárias não correspondem a nenhuma instituição
> real e os resultados obtidos com ela **não têm validade regulatória**.

---

## 📥 Download para teste (sem instalar nada)

Baixe o **`Auditoria_Matrizes.zip`** na página de
**[Releases](../../releases/latest)**, extraia a pasta inteira e execute
`Auditoria_Matrizes.exe` (Windows, não requer Python).

> Na primeira execução o Windows pode exibir o aviso do SmartScreen — clique em
> **"Mais informações" → "Executar assim mesmo"** (o executável não possui
> assinatura digital). Instruções completas no `LEIA-ME.txt` dentro do pacote.

## 🖥️ O que a ferramenta faz

1. Lê uma base CSV com as matrizes curriculares (uma linha por disciplina, com a
   carga horária fragmentada em teórica, prática de sala, prática de laboratório,
   EAD, ASM, estágio e extensão);
2. Valida cada matriz com o CP-SAT contra as regras regulatórias:

| Regra | Limite | Escopo |
|---|---|---|
| Extensão | ≥ 10% da CH total | todos os cursos |
| EAD | ≤ 40% da CH | cursos presenciais |
| ASM | ≤ 20% da CH | cursos que possuem ASM |
| Estágio | Medicina ≥ 35% · demais Saúde ≥ 20% · Engenharias ≥ 160h · Licenciaturas ≥ 400h | por área/curso |
| Prática pedagógica | ≥ 400h | Licenciaturas |
| Prática indevida | CH de prática = 0 | cursos cuja DCN não prevê prática (Economia, Ciência da Computação, Sistemas de Informação) |
| Integridade | soma dos componentes = 100% da CH | todas as matrizes |

3. Exibe o passo a passo da análise em tempo real na interface;
4. Gera o relatório `Relatorio_Auditoria_Matrizes.xlsx` com duas abas:
   **Resumo Executivo** (todas as matrizes) e **Não Conformidades** (matrizes
   reprovadas, com o status de cada regra e o parecer técnico).

## 🚀 Executando a partir do código-fonte

Requisitos: Python 3.9+

```bash
pip install -r requirements.txt
python INTERFACE.py      # interface gráfica
python PLANEJADOR.PY     # execução direta no terminal (CLI)
```

Para gerar o executável distribuível:

```bash
pip install pyinstaller
python build_exe.py      # saída em dist/Auditoria_Matrizes/
```

## 📁 Estrutura do repositório

| Arquivo | Descrição |
|---|---|
| `PLANEJADOR.PY` | Motor de auditoria: ingestão do CSV, validação CP-SAT e relatório Excel |
| `INTERFACE.py` | Interface gráfica Tkinter (log ao vivo, barra de progresso, KPIs) |
| `build_exe.py` | Empacotamento com PyInstaller (modo pasta, abertura instantânea) |
| `Dados_Exportacao_CSV.csv` | Base **fictícia** de exemplo — 60 matrizes, 23 cursos, 5 áreas (serve de template) |
| `LEIA-ME.txt` | Guia do usuário final (acompanha o executável) |
| `PLANEJADOR_ARTIGO.txt` | Artigo didático: como aplicar OR-Tools/CP-SAT a qualquer problema |
| `requirements.txt` | Dependências Python |

## 📄 Formato do CSV de entrada

Separador `;`, uma linha por disciplina. Colunas: `ID_Matriz`, `Area`, `Curso`,
`Modalidade`, `ID_Disciplina`, `Nome_Disciplina`, `Sinergia`, `Carga_Horaria`,
`CH_Teorica`, `CH_Pratica_Sala`, `CH_Pratica_Lab`, `CH_EAD`, `CH_ASM`,
`CH_Estagio`, `CH_Extensao`, `Eh_EAD`. Use a base de exemplo como template:
mantenha o cabeçalho e substitua as linhas pelos seus dados.

## 🧠 Quer entender (e reaproveitar) a técnica?

Leia o **[`PLANEJADOR_ARTIGO.txt`](PLANEJADOR_ARTIGO.txt)** — um guia didático
que explica o método em 5 passos (descrever o problema → declarar variáveis →
restrições → função objetivo → resolver), como este caso usa o "truque das
variáveis travadas" para transformar o solver em auditor, e como construir a
interface Tkinter sem congelar a janela. O mesmo esqueleto serve para escalas
de trabalho, distribuição de orçamento, grades de horário e qualquer problema
de restrições.

## ⚖️ Aviso legal

Os limites regulatórios implementados refletem leitura das resoluções vigentes
em julho/2026 (referências no artigo e nos comentários do código). Esta
ferramenta é um apoio à análise — **não substitui** a avaliação da procuradoria
institucional ou dos órgãos reguladores.
