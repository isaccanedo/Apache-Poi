## Apache POI

Este módulo contém artigos sobre o Microsoft Word Processing in Java with Apache POI

# 1. Visão Geral
Apache POI é uma biblioteca Java para trabalhar com os vários formatos de arquivo baseados nos padrões Office Open XML (OOXML) e no formato OLE 2 Compound Document da Microsoft (OLE2).

Este tutorial se concentra no suporte do Apache POI para Microsoft Word, o formato de arquivo do Office mais comumente usado. Ele percorre as etapas necessárias para formatar e gerar um arquivo MS Word e como analisá-lo.

# 2. Dependências Maven
A única dependência necessária para o Apache POI lidar com arquivos do MS Word é:
```
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>3.15</version>
</dependency
```

Clique aqui para obter a versão mais recente deste artefato.
https://search.maven.org/classic/#search%7Cga%7C1%7Cg%3A%22org.apache.poi%22%20AND%20a%3A%22poi-ooxml%22

# 3. Preparação
Vejamos agora alguns dos elementos usados para facilitar a geração de um arquivo MS Word.

### 3.1. Arquivos de recursos
Vamos coletar o conteúdo de três arquivos de texto e gravá-los em um arquivo MS Word - denominado rest-with-spring.docx.

Além disso, o arquivo logo-leaf.png é usado para inserir uma imagem nesse novo arquivo. Todos esses arquivos existem no caminho de classe e são representados por várias variáveis estáticas:
```
public static String logo = "logo-leaf.png";
public static String paragraph1 = "poi-word-para1.txt";
public static String paragraph2 = "poi-word-para2.txt";
public static String paragraph3 = "poi-word-para3.txt";
public static String output = "rest-with-spring.docx";
```

Para os curiosos, os conteúdos desses arquivos de recursos do repositório, cujo link é dado na última seção deste tutorial, são extraídos desta página do curso aqui no site.

### 3.2. Método Auxiliar
O método principal que consiste na lógica usada para gerar um arquivo MS Word, que é descrito na seção a seguir, faz uso de um método auxiliar:

```
public String convertTextFileToString(String fileName) {
    try (Stream<String> stream 
      = Files.lines(Paths.get(ClassLoader.getSystemResource(fileName).toURI()))) {
        
        return stream.collect(Collectors.joining(" "));
    } catch (IOException | URISyntaxException e) {
        return null;
    }
}
```

Este método extrai o conteúdo contido em um arquivo de texto localizado no caminho de classe, cujo nome é o argumento String transmitido. Em seguida, ele concatena as linhas neste arquivo e retorna a String de junção.

# 4. Geração de arquivo MS Word
Esta seção fornece instruções sobre como formatar e gerar um arquivo do Microsoft Word. Antes de trabalhar em qualquer parte do arquivo, precisamos ter uma instância XWPFDocument:

```
XWPFDocument document = new XWPFDocument();
```

### 4.1. Formatando Título e Subtítulo
Para criar o título, precisamos primeiro instanciar a classe XWPFParagraph e definir o alinhamento no novo objeto:
```
XWPFParagraph title = document.createParagraph();
title.setAlignment(ParagraphAlignment.CENTER);
```

O conteúdo de um parágrafo precisa ser agrupado em um objeto XWPFRun. Podemos configurar este objeto para definir um valor de texto e seus estilos associados:

```
XWPFRun titleRun = title.createRun();
titleRun.setText("Build Your REST API with Spring");
titleRun.setColor("009933");
titleRun.setBold(true);
titleRun.setFontFamily("Courier");
titleRun.setFontSize(20);
```

Deve-se ser capaz de inferir os objetivos dos métodos de conjunto a partir de seus nomes.

De maneira semelhante, criamos uma instância XWPFParagraph que inclui o subtítulo:

```
XWPFParagraph subTitle = document.createParagraph();
subTitle.setAlignment(ParagraphAlignment.CENTER);
```

Vamos formatar o subtítulo também:

```
XWPFRun subTitleRun = subTitle.createRun();
subTitleRun.setText("from HTTP fundamentals to API Mastery");
subTitleRun.setColor("00CC44");
subTitleRun.setFontFamily("Courier");
subTitleRun.setFontSize(16);
subTitleRun.setTextPosition(20);
subTitleRun.setUnderline(UnderlinePatterns.DOT_DOT_DASH);
```

O método setTextPosition define a distância entre a legenda e a imagem subsequente, enquanto setUnderline determina o padrão de sublinhado.

Observe que codificamos permanentemente o conteúdo do título e do subtítulo, pois essas instruções são muito curtas para justificar o uso de um método auxiliar.

### 4.2. Inserindo uma Imagem

Uma imagem também precisa ser agrupada em uma instância XWPFParagraph. Queremos que a imagem seja centralizada horizontalmente e colocada sob o subtítulo, portanto, o seguinte trecho deve ser colocado abaixo do código fornecido acima:

```
XWPFParagraph image = document.createParagraph();
image.setAlignment(ParagraphAlignment.CENTER);
```

Veja como definir a distância entre esta imagem e o texto abaixo dela:

```
XWPFRun imageRun = image.createRun();
imageRun.setTextPosition(20);
```

Uma imagem é obtida de um arquivo no caminho de classe e, em seguida, inserida no arquivo MS Word com as dimensões especificadas:

```
Path imagePath = Paths.get(ClassLoader.getSystemResource(logo).toURI());
imageRun.addPicture(Files.newInputStream(imagePath),
  XWPFDocument.PICTURE_TYPE_PNG, imagePath.getFileName().toString(),
  Units.toEMU(50), Units.toEMU(50));
```

### 4.3. Formatando Parágrafos
Aqui está como criamos o primeiro parágrafo com conteúdo retirado do arquivo poi-word-para1.txt:

```
XWPFParagraph para1 = document.createParagraph();
para1.setAlignment(ParagraphAlignment.BOTH);
String string1 = convertTextFileToString(paragraph1);
XWPFRun para1Run = para1.createRun();
para1Run.setText(string1);
```

É evidente que a criação de um parágrafo é semelhante à criação do título ou subtítulo. A única diferença aqui é o uso do método auxiliar em vez de strings embutidas em código.

De maneira semelhante, podemos criar dois outros parágrafos usando o conteúdo dos arquivos poi-word-para2.txt e poi-word-para3.txt:

```
XWPFParagraph para2 = document.createParagraph();
para2.setAlignment(ParagraphAlignment.RIGHT);
String string2 = convertTextFileToString(paragraph2);
XWPFRun para2Run = para2.createRun();
para2Run.setText(string2);
para2Run.setItalic(true);

XWPFParagraph para3 = document.createParagraph();
para3.setAlignment(ParagraphAlignment.LEFT);
String string3 = convertTextFileToString(paragraph3);
XWPFRun para3Run = para3.createRun();
para3Run.setText(string3);
```

A criação desses três parágrafos é quase a mesma, exceto por alguns estilos, como alinhamento ou itálico.

### 4.4. Gerando arquivo MS Word
Agora estamos prontos para gravar um arquivo do Microsoft Word na memória a partir da variável do documento:

```
FileOutputStream out = new FileOutputStream(output);
document.write(out);
out.close();
document.close();
```

Todos os fragmentos de código nesta seção são agrupados em um método denominado handleSimpleDoc.

# 5. Análise e teste

Esta seção descreve a análise de arquivos do MS Word e a verificação do resultado.

### 5.1. Preparação
Declaramos um campo estático na classe de teste:

```
static WordDocument wordDocument;
```

Este campo é usado para fazer referência a uma instância da classe que inclui todos os fragmentos de código mostrados nas seções 3 e 4.

Antes de analisar e testar, precisamos inicializar a variável estática declarada logo acima e gerar o arquivo rest-with-spring.docx no diretório de trabalho atual invocando o método handleSimpleDoc:

```
@BeforeClass
public static void generateMSWordFile() throws Exception {
    WordTest.wordDocument = new WordDocument();
    wordDocument.handleSimpleDoc();
}
```

Vamos passar para a etapa final: analisar o arquivo MS Word e verificar o resultado.

### 5.2 Analisando Arquivo do MS Word e Verificação
Primeiro, extraímos o conteúdo do arquivo MS Word fornecido no diretório do projeto e armazenamos o conteúdo em uma Lista de XWPFParagraph:

```
Path msWordPath = Paths.get(WordDocument.output);
XWPFDocument document = new XWPFDocument(Files.newInputStream(msWordPath));
List<XWPFParagraph> paragraphs = document.getParagraphs();
document.close();
```

A seguir, vamos nos certificar de que o conteúdo e o estilo do título sejam iguais aos que definimos antes:

```
XWPFParagraph title = paragraphs.get(0);
XWPFRun titleRun = title.getRuns().get(0);
 
assertEquals("Build Your REST API with Spring", title.getText());
assertEquals("009933", titleRun.getColor());
assertTrue(titleRun.isBold());
assertEquals("Courier", titleRun.getFontFamily());
assertEquals(20, titleRun.getFontSize());
```

Por uma questão de simplicidade, apenas validamos o conteúdo de outras partes do arquivo, deixando de fora os estilos. A verificação de seus estilos é semelhante ao que fizemos com o título:

```
assertEquals("from HTTP fundamentals to API Mastery",
  paragraphs.get(1).getText());
assertEquals("What makes a good API?", paragraphs.get(3).getText());
assertEquals(wordDocument.convertTextFileToString
  (WordDocument.paragraph1), paragraphs.get(4).getText());
assertEquals(wordDocument.convertTextFileToString
  (WordDocument.paragraph2), paragraphs.get(5).getText());
assertEquals(wordDocument.convertTextFileToString
  (WordDocument.paragraph3), paragraphs.get(6).getText());
```

Agora podemos ter certeza de que a criação do arquivo rest-with-spring.docx foi bem-sucedida.#

# 6. Conclusão 

Este tutorial introduziu o suporte do Apache POI para o formato Microsoft Word. Ele passou por etapas necessárias para gerar um arquivo do MS Word e verificar seu conteúdo.