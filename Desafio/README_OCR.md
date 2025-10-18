# Processamento de PDFs com OCR

Este projeto agora suporta processamento de PDFs escaneados ou com imagens usando OCR (Optical Character Recognition).

## Instalação do Tesseract OCR

Para usar a funcionalidade de OCR, você precisa instalar o Tesseract OCR:

### Windows

1. **Download**: Baixe o instalador do Tesseract em:
   https://github.com/UB-Mannheim/tesseract/wiki

2. **Instalação**: Execute o instalador e instale em:
   `C:\Program Files\Tesseract-OCR\`

3. **Verificação**: Após a instalação, você deve ter:
   - `C:\Program Files\Tesseract-OCR\tesseract.exe`
   - `C:\Program Files\Tesseract-OCR\tessdata\` (pasta com dados de treinamento)

### Linux (Ubuntu/Debian)

```bash
sudo apt-get update
sudo apt-get install tesseract-ocr tesseract-ocr-por
```

### macOS

```bash
brew install tesseract tesseract-lang
```

## Como Usar

### Interface Gráfica (Windows Forms)

1. Execute o programa sem argumentos: `dotnet run`
2. Selecione um arquivo PDF
3. Marque a opção "Usar OCR (para PDFs escaneados/imagens)"
4. Clique em "Processar Documento"

### Linha de Comando

```bash
# Processamento normal (texto)
dotnet run -- Exemplos/Exemplo-Cartao-Ponto-01.pdf saida_normal.xlsx

# Processamento com OCR (imagens)
dotnet run -- Exemplos/Exemplo-Cartao-Ponto-01.pdf saida_ocr.xlsx --ocr
```

## Funcionalidades do OCR

### Cartões de Ponto
- Extrai horários de entrada/saída
- Identifica dias da semana
- Processa intervalos
- Detecta situação (S/N)
- Calcula horas trabalhadas

### Holerites
- Extrai códigos de proventos/descontos
- Processa descrições
- Calcula valores e quantidades
- Identifica totais

## Limitações

- **Qualidade da imagem**: OCR funciona melhor com imagens de alta qualidade
- **Fonte**: Textos manuscritos são mais difíceis de processar
- **Idioma**: Configurado para português brasileiro
- **Performance**: OCR é mais lento que processamento de texto

## Troubleshooting

### Erro: "Tesseract not found"
- Verifique se o Tesseract está instalado
- Confirme o caminho da instalação
- Reinicie o Visual Studio/IDE

### Baixa precisão do OCR
- Use imagens com resolução maior
- Certifique-se de que o texto está legível
- Evite documentos muito escuros ou claros

### Erro de memória
- Processe um PDF por vez
- Feche outros programas se necessário

## Exemplo de Uso

```csharp
// Criar serviço OCR
var ocrService = new ImageBasedPdfService();

// Processar cartão de ponto
var timeCardData = ocrService.ProcessTimeCardWithOCR("cartao_escaneado.pdf");

// Processar holerite
var payrollData = ocrService.ProcessPayrollWithOCR("holerite_escaneado.pdf");
```

## Dependências Adicionais

- `Tesseract` (5.2.0) - Engine OCR
- `SixLabors.ImageSharp` (3.0.2) - Processamento de imagens
- `PdfPig` (0.1.11) - Extração de imagens do PDF
