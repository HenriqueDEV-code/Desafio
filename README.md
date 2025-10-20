# 🚀 Quick Filler - Processador de Documentos PDF

## 📋 Descrição

O **Processador de Documentos** é uma aplicação desktop desenvolvida em C# que automatiza o processamento de documentos PDF relacionados à folha de pagamento, especificamente:

- **📊 Holerites** - Extrai proventos, descontos e informações de pagamento
- **⏰ Cartões de Ponto** - Processa horários de entrada, saída e horas trabalhadas

A aplicação converte automaticamente esses documentos PDF em planilhas Excel (.xlsx) estruturadas e organizadas, facilitando análises e cálculos.

## ✨ Funcionalidades

### 🎯 Processamento Inteligente
- **Detecção Automática**: Identifica automaticamente o tipo de documento (holerite ou cartão de ponto)


### 📊 Tipos de Documentos Suportados

#### 💰 Holerites
- Extrai informações de proventos (salário, horas extras, bonificações)
- Processa descontos (INSS, IRRF, vale transporte, etc.)
- Identifica dados do funcionário e período de pagamento
- Calcula valores líquidos e brutos

#### ⏰ Cartões de Ponto
- Processa horários de entrada e saída diários
- Calcula horas trabalhadas por dia
- Identifica períodos de intervalo
- Extrai informações de horas extras

### 🖥️ Interface de Usuário
- **Interface Gráfica**: Aplicação Windows Forms moderna e intuitiva
- **Modo Linha de Comando**: Suporte para automação e scripts
- **Barra de Progresso**: Acompanhamento visual do processamento
- **Seleção de Arquivos**: Interface amigável para escolha de documentos

## 🛠️ Tecnologias e Dependências

### 📦 Framework Principal
- **.NET 8.0 Desktop Runtime** (Obrigatório)
- **Windows Forms** para interface gráfica
- **C# 12** com recursos modernos

### 📚 Bibliotecas NuGet Utilizadas

| Biblioteca | Versão | Propósito |
|------------|--------|-----------|
| **ClosedXML** | 0.105.0 | Geração de arquivos Excel (.xlsx) |
| **PdfiumViewer** | 2.13.0 | Visualização e manipulação de PDFs |
| **PdfPig** | 0.1.11 | Extração de texto e dados de PDFs |
| **Tesseract** | 5.2.0 | Motor OCR para reconhecimento de texto |
| **TesseractOCR** | 5.5.1 | Wrapper .NET para Tesseract |
| **SixLabors.ImageSharp** | 3.1.11 | Processamento de imagens |
| **System.Text.Json** | 8.0.5 | Serialização JSON |

### 🔧 Dependências do Sistema

#### ⚠️ Pré-requisitos Obrigatórios
1. **.NET Desktop Runtime 8.0** (Microsoft.WindowsDesktop.App x64)
   - Download: [Microsoft .NET 8.0 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0)
   - **IMPORTANTE**: Sem esta dependência, o instalador não funcionará

2. **Windows 10/11** (64-bit)
3. **Visual C++ Redistributable** (incluído automaticamente)

## 📁 Estrutura do Projeto

```
├── 📁 .git/ 🚫 (auto-hidden)
├── 📁 .vs/ 🚫 (auto-hidden)
├── 📁 Desafio/
│   ├── 📁 Controllers/
│   │   └── 🟣 DocumentProcessor.cs      # Controlador principal
│   ├── 📁 Instalador/
│   │   └── ⚙️ Setup_Desafio - Teste.exe # Instalador da aplicação
│   ├── 📁 Models/
│   │   ├── 🟣 PayrollData.cs           # Modelo de dados para holerites
│   │   └── 🟣 TimeCardData.cs          # Modelo de dados para cartões de ponto
│   ├── 📁 Services/
│   │   ├── 🟣 ExcelGenerator.cs        # Geração de planilhas Excel
│   │   ├── 🟣 ImageBasedPdfService.cs # Processamento com OCR
│   │   ├── 🟣 PayrollPdfService.cs    # Processamento de holerites
│   │   └── 🟣 TimeCardPdfService.cs   # Processamento de cartões de ponto
│   ├── 📁 bin/ 🚫 (auto-hidden)
│   ├── 📁 obj/ 🚫 (auto-hidden)
│   ├── 🟣 Desafio.csproj              # Arquivo de projeto
│   ├── 📄 Desafio.csproj.user 🚫 (auto-hidden)
│   ├── 🟣 MainForm.cs                 # Interface principal
│   ├── 🟣 Program.cs                  # Ponto de entrada da aplicação
│   └── 📖 README.md                   # Este arquivo
├── 📁 Exemplos/
│   ├── 📕 Exemplo-Cartao-Ponto-01.pdf # Exemplo de cartão de ponto
│   ├── 📕 Exemplo-Holerite-01.pdf     # Exemplo de holerite
│   ├── 📕 Exemplo-Holerite-02.pdf     # Exemplo de holerite escaneado
│   └── 📊 Exemplo-Holerite-02_ocr.xlsx # Resultado do processamento OCR
├── 📄 .gitattributes
├── 🚫 .gitignore
├── 🟣 Desafio.sln                     # Solução Visual Studio
└── 📜 LICENSE.txt                     # Licença do projeto
```

## 🚀 Instalação

### 📥 Processo de Instalação

#### 1️⃣ **Pré-requisito: Instalar .NET Desktop Runtime 8.0**

Antes de executar o instalador, você **DEVE** instalar o .NET Desktop Runtime 8.0:

![Pré-requisito .NET](https://via.placeholder.com/400x200/FFD700/000000?text=.NET+Desktop+Runtime+8.0+Required)

- **Download**: [Microsoft .NET 8.0 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0)
- **Versão**: Microsoft.WindowsDesktop.App x64
- **Importante**: Sem esta dependência, você verá o erro mostrado na imagem acima

#### 2️⃣ **Executar o Instalador**

1. Navegue até a pasta `Instalador/`
2. Execute o arquivo `Setup_Desafio - Teste.exe`
3. Siga o assistente de instalação

#### 3️⃣ **Configurações do Instalador**

![Configurações do Instalador](https://via.placeholder.com/500x400/4CAF50/FFFFFF?text=Installer+Configuration)

O instalador oferece as seguintes opções:
- ✅ **Criar atalho na Área de Trabalho** (recomendado)
- ✅ **Executar Desafio - Teste após concluir a instalação** (recomendado)

#### 4️⃣ **Processo de Instalação**

![Instalação em Progresso](https://via.placeholder.com/500x300/2196F3/FFFFFF?text=Installation+Progress)

Durante a instalação, você verá:
- Barra de progresso indicando o status
- Arquivos sendo extraídos (incluindo bibliotecas OCR)
- Instalação das dependências necessárias

#### 5️⃣ **Conclusão da Instalação**

![Instalação Concluída](https://via.placeholder.com/500x400/8BC34A/FFFFFF?text=Installation+Complete)

Após a instalação bem-sucedida:
- ✅ Aplicação instalada em `C:\Program Files\Desafio - Teste\`
- ✅ Atalho criado na área de trabalho (se selecionado)
- ✅ Aplicação pode ser iniciada pelos atalhos instalados

### 📍 Localização dos Arquivos

Após a instalação, os arquivos ficam localizados em:

```
C:\Program Files\Desafio - Teste\
├── 📁 x64/                    # Bibliotecas 64-bit
├── 📁 x86/                    # Bibliotecas 32-bit
├── 📁 nl/                     # Recursos de idioma
├── 🟣 Desafio.exe            # Executável principal
├── 📚 *.dll                  # Bibliotecas de dependência
└── 🗑️ unins000.exe           # Desinstalador
```

### 🗑️ Desinstalação

Para desinstalar a aplicação:
1. Execute `C:\Program Files\Desafio - Teste\unins000.exe`
2. Ou use "Adicionar ou Remover Programas" no Windows

## 💻 Como Usar

### 🖥️ Interface Gráfica

1. **Iniciar a Aplicação**
   - Clique no atalho da área de trabalho
   - Ou execute `Desafio.exe` diretamente

2. **Selecionar Arquivo PDF**
   - Clique em "Procurar..." para escolher o documento
   - A aplicação sugere arquivos da pasta `Exemplos/`

3. **Configurar Processamento**
   - ✅ Marque "Usar OCR" para PDFs escaneados/imagens
   - ✅ Deixe desmarcado para PDFs com texto selecionável

4. **Processar Documento**
   - Clique em "Processar Documento"
   - Acompanhe o progresso na barra de status
   - O arquivo Excel será salvo na mesma pasta do PDF

### ⌨️ Linha de Comando

```bash
# Sintaxe básica
Desafio.exe <caminho_pdf> <caminho_saida>

# Exemplos
Desafio.exe "Exemplos\Exemplo-Cartao-Ponto-01.pdf" "cartao_ponto_saida.xlsx"
Desafio.exe "Exemplos\Exemplo-Holerite-01.pdf" "holerite_saida.xlsx"
```

## 📊 Exemplos de Uso

### 📄 Arquivos de Exemplo Incluídos

A pasta `Exemplos/` contém documentos de teste:

- **`Exemplo-Cartao-Ponto-01.pdf`** - Cartão de ponto padrão
- **`Exemplo-Holerite-01.pdf`** - Holerite com texto selecionável
- **`Exemplo-Holerite-02.pdf`** - Holerite escaneado (requer OCR)
- **`Exemplo-Holerite-02_ocr.xlsx`** - Resultado do processamento OCR

### 🎯 Resultados Esperados

#### 📊 Holerite Processado
- Planilha com abas organizadas por categoria
- Proventos detalhados (salário, horas extras, etc.)
- Descontos listados (INSS, IRRF, vale transporte)
- Totais calculados automaticamente

#### ⏰ Cartão de Ponto Processado
- Planilha com dias da semana organizados
- Horários de entrada e saída por dia
- Cálculo automático de horas trabalhadas
- Identificação de horas extras

## ⚠️ Limitações e status da extração por imagem (OCR)

Atualmente, a extração baseada em imagem para processar PDFs escaneados e gerar informações estruturadas no Excel **não funcionou de forma confiável** neste projeto, mesmo utilizando as bibliotecas previstas. Abaixo estão os motivos principais:

- Problema de dados de idioma (tessdata): o Tesseract necessita dos arquivos de idioma (por exemplo, `por.traineddata` para português e `osd.traineddata` para detecção de orientação). Esses arquivos não estão incluídos por padrão com os .dlls e, se ausentes ou incorretos, o OCR retorna texto vazio ou com baixa acurácia.
- Qualidade do PDF escaneado: imagens com baixa resolução, compressão alta, desalinhamento (skew), ruído e sombras reduzem drasticamente a qualidade do OCR; sem um pré-processamento robusto (binarização, deskew, de-noise, contraste), o reconhecimento falha.
- Layouts complexos de holerites/cartões: holerites e cartões de ponto possuem tabelas e colunas; o OCR entrega texto sem estrutura. Sem uma camada de reconstrução de tabelas (detecção de linhas/células) e regras específicas por layout, não é possível mapear com confiabilidade os valores para o Excel.
- Limitações das bibliotecas no contexto atual: a combinação atual (PdfPig para extração/varredura de páginas, ImageSharp para pré-processamento simples e Tesseract para OCR) não inclui um pipeline completo de detecção de tabelas e pós-processamento semântico, o que é necessário para transformar texto OCR em dados estruturados.
- PDFs híbridos ou protegidos: alguns PDFs misturam texto e imagem ou possuem proteções/formatações internas que dificultam a extração de imagens em qualidade adequada para OCR.

Em resumo: o OCR foi mantido como funcionalidade experimental e pode não extrair informações suficientes para preencher corretamente o Excel, especialmente em documentos escaneados com qualidade mediana/baixa ou com layouts não previstos.

### O que já está implementado
- Detecção de tipo de documento por texto (quando o PDF contém texto selecionável).
- Pipeline de OCR básico para tentar extrair texto de imagens.
- Geração de planilha Excel a partir de dados quando a extração é bem-sucedida.

### Caminhos recomendados para melhorar a extração por imagem
- Incluir e distribuir `tessdata` com `por.traineddata` e `osd.traineddata` adequados.
- Adicionar pré-processamento de imagem mais robusto (deskew, binarização adaptativa, remoção de ruído, aumento de contraste, corte de bordas).
- Implementar detecção e reconstrução de tabelas (detecção de linhas/células) antes do parsing semântico.
- Criar parsers específicos por layout (regras/regex por fornecedor/modelo de holerite/cartão).
- Considerar serviços OCR com melhor acurácia e layout analysis (Azure OCR, Google Cloud Vision, ABBYY) se necessário.

## 🔧 Desenvolvimento

### 🏗️ Compilação

```bash
# Restaurar dependências
dotnet restore

# Compilar projeto
dotnet build

# Executar aplicação
dotnet run
```


## 📝 Licença

Este projeto está licenciado sob os termos especificados no arquivo `LICENSE.txt`.

---

**Desenvolvido com ❤️ em C# e .NET 8.0**
