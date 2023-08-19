# Funciones para Justificar Texto en SSRS con Tipos de Fuentes Monoespaciadas

Este archivo README proporciona una descripción de las funciones en Visual Basic utilizadas para justificar texto en informes de SQL Server Reporting Services (SSRS) con tipos de fuentes monoespaciadas.

## Índice

- [Descripción General](#descripción-general)
    - [Función TextOnlyJustify](#función-textonlyjustify)
        - [Función TextLine](#función-textline)
        - [Función ReplaceSome](#función-replacesome)
    - [Función TextOnlyJustifyHTML](#función-textonlyjustifyhtml)
    - [Función TextOnlyJustifyHtmlBold](#función-textonlyjustifyhtmlbold)
        - [Función ApplyBoldFormat](#función-applyboldformat)
- [Cómo Utilizar las Funciones](#cómo-utilizar-las-funciones)

## Descripción General

El conjunto de funciones proporcionado en este archivo tiene como objetivo permitir la justificación de texto en informes de SSRS cuando se utilizan fuentes monoespaciadas. Estas fuentes aseguran que cada carácter ocupa el mismo ancho, lo que facilita el proceso de alineación del texto.

Las funciones principales son:

### Función TextOnlyJustify

Esta función toma un bloque de texto, una fuente específica, un valor booleano para la indentación y un ancho máximo para las líneas. Justifica el texto completo dividiéndolo en párrafos y palabras, y luego aplicando justificación a las líneas individuales según la fuente y el ancho proporcionados. Devuelve el texto justificado como una cadena.

```vb
Public Function TextOnlyJustify(text As String, font As System.Drawing.Font, bIndent As Boolean, width As Single) As String
    Dim bmp As New System.Drawing.Bitmap(1024, 1024)
    Dim gr = System.Drawing.Graphics.FromImage(bmp)
    Dim sRtn As New System.Text.StringBuilder()
 
    Dim paragraphs As String() = text.Split(ControlChars.NewLine)
 
    For Each paragraph As String In paragraphs
        Dim words As String() = paragraph.Split(" "C)
        Dim start_word As Integer = 0
        Dim indent As Single = If(bIndent, 40F, 0F)
 
        ' Repeat until we run out of text or room.
        While True
            ' See how many words will fit.
            ' Start with just the next word.
            Dim line As String = words(start_word)
 
            ' Add more words until the line won't fit.
            Dim end_word As Integer = start_word + 1
            While end_word < words.Length
                ' See if the next word fits.
                Dim test_line As String = (line & Convert.ToString(" ")) + words(end_word)
                Dim line_size As System.Drawing.SizeF = gr.MeasureString(test_line, font)
                If line_size.Width + indent > width Then
                    ' The line is too wide. Don't use the last word.
                    end_word -= 1
                    Exit While
                Else
                    ' The word fits. Save the test line.
                    line = test_line
                End If
 
                ' Try the next word.
                end_word += 1
            End While
 
            ' See if this is the last line in the paragraph.
            If (end_word = words.Length) Then
                ' This is the last line. Don't justify it.
                sRtn.Append(line)
            Else
                ' This is not the last line. Justify it.
 
                sRtn.Append(TextLine(gr, line, font, width - indent, True))
            End If
 
            ' Start the next line at the next word.
            start_word = end_word + 1
            If start_word >= words.Length Then
                Exit While
            End If
 
            ' Don't indent subsequent lines in this paragraph.
            indent = 0
        End While
 
        ' Add a gap after the paragraph.
        sRtn.Append(vbLf & vbCr)
    Next
 
    Return sRtn.ToString()
End Function
```

 - text: El texto que se desea justificar.
 - font: La fuente utilizada para el texto.
 - bIndent: Un valor booleano que determina si se debe indentar el texto.
 - width: El ancho máximo al que se ajustará el texto justificado.

Ejemplo de uso :
```vb
=Code.TextOnlyJustify(
"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Maecenas ac lorem nec diam auctor aliquam sit amet quis augue. Praesent pharetra dolor malesuada eros laoreet volutpat. Nulla ac rhoncus lorem. Pellentesque vel tempus ex, ut sagittis eros. Curabitur rutrum lacus risus, vel molestie augue euismod quis. Nunc dapibus convallis lectus, eu mattis nisi dictum eget. Morbi quis enim eget dolor ultricies ultrices id in lacus. Sed nec tellus ut enim lacinia malesuada. Etiam sit amet sollicitudin erat." +
"Curabitur auctor hendrerit pulvinar. Ut imperdiet arcu vitae rutrum iaculis. Donec at est cursus, laoreet turpis quis, eleifend odio. Donec at mattis magna. Praesent odio ligula, tincidunt et erat non, laoreet semper metus. Praesent hendrerit elit quis commodo commodo. Pellentesque non porttitor ipsum. Aenean eu mauris ut lorem varius tincidunt sed id dui. Proin euismod nec orci quis rutrum. Nunc gravida gravida posuere. Suspendisse nec porttitor orci. Donec consectetur nisl eu tempor feugiat. Cras vulputate sem quis mauris mattis, nec tincidunt orci pulvinar. Proin viverra ipsum tellus, quis accumsan magna fringilla sit amet. Aliquam egestas augue nec lacus suscipit sagittis. Nulla quis sodales lorem."
,
new System.Drawing.Font("Consolas", "6"), 
true, 
715)
```

### Función TextLine

Esta función realiza la justificación de una línea de texto individual. Toma una instancia de `System.Drawing.Graphics`, la línea de texto, la fuente, el ancho máximo y un valor booleano que indica si se debe aplicar justificación. Esta función es utilizada por `TextOnlyJustify` para justificar líneas individuales.

```vb
Public Function TextLine(gr As System.Drawing.Graphics, line As String, font As System.Drawing.Font, width As Single, justification As Boolean) As String
    Dim sLine As New System.Text.StringBuilder()
    ' See if we should use full justification.
    If justification Then
        ' Justify the text.
        ' Break the text into words.
        Dim words As String() = line.Split(" "c)
 
        ' Add a space to each word and get their lengths.
        Dim word_width As Single() = New Single(words.Length - 1) {}
        Dim total_width As Single = 0
        For i As Integer = 0 To words.Length - 1
            ' See how wide this word is.
            Dim size As System.Drawing.SizeF = gr.MeasureString(words(i), font)
            word_width(i) = size.Width
            total_width += word_width(i)
        Next
 
        ' Get the additional spacing between words.
        Dim extra_space As Single = width - total_width
        Dim num_spaces As Integer = words.Length - 1
        If words.Length > 1 Then
            extra_space /= (num_spaces-1)
        End If
 
        For i2 As Integer = 1 To 100
            Dim sTest As String = line.Replace(" ", New String(ChrW(&H200A), i2))
            If gr.MeasureString(sTest, font).Width > width Then
 
                For i3 As Integer = words.Length To 1 Step -1
                    sTest = line.Replace(" ", New String(ChrW(&H200A), i2 - 1))
                    Dim sTemp = ReplaceSome(sTest, New String(ChrW(&H200A), i2 - 1), New String(ChrW(&H200A), i2), i3)
                    If gr.MeasureString(sTemp, font).Width < width Then
                        Console.WriteLine("{0}, size: {1}", line, gr.MeasureString(sTemp, font).Width)
                        Return sTemp + ControlChars.CrLf
                    End If
                Next
 
                Console.WriteLine("{0}, size: {1}", line, gr.MeasureString(line.Replace(" ", New String(ChrW(&H200A), i2 - 1)), font).Width)
                Return line.Replace(" ", New String(ChrW(&H200A), i2 - 1)) + ControlChars.CrLf
            End If
        Next
 
 
    Else
        Return line
    End If
End Function
```

 - gr: Un objeto Graphics que permite manipular elementos gráficos.
 - line: La línea de texto a justificar.
 - font: La fuente utilizada para el texto.
 - width: El ancho máximo al que se ajustará el texto.
 - justification: Un valor booleano que determina si se debe justificar el texto.

Si justification es verdadero, la función justificará el texto dentro del ancho especificado, dividiendo las palabras y añadiendo espacios adicionales entre ellas para que se ajusten. Si no se necesita justificación, la función devuelve la línea de texto original.

### Función ReplaceSome

Esta función auxiliar reemplaza una parte específica de una cadena con otra. Se utiliza para ajustar el espaciado entre palabras en la línea justificada.

```vb
Private Function ReplaceSome(s As String, repl As String, wth As String, num As Integer) As String
 
    ReplaceSome = String.Empty
    Dim s2 As String() = s.Split(repl, num, StringSplitOptions.RemoveEmptyEntries)
 
    For t As Integer = 0 To s2.Length - 2
        ReplaceSome += s2(t) + wth
    Next
 
    ReplaceSome += s2(s2.Length - 1)
End Function
```

 - s: La cadena original en la que se realizarán los reemplazos.
 - repl: La subcadena que se reemplazará.
 - wth: La subcadena que se utilizará como reemplazo.
 - num: El número máximo de reemplazos que se realizarán.

La función divide la cadena original en subcadenas usando repl como delimitador. Luego, reemplaza solo un número específico de esas subcadenas con wth, y finalmente vuelve a unir todas las subcadenas modificadas.

### Función TextOnlyJustifyHTML

Esta función es similar a `TextOnlyJustify`, pero en lugar de devolver una cadena de texto, devuelve una versión en formato HTML con etiquetas `<br />` para los saltos de línea y espacios no rompibles (`&nbsp;`) para el espaciado adicional entre palabras.

```vb
Public Function TextOnlyJustifyHTML(text As String, maxCharsPerLine As Integer) As String
    Dim words() As String = text.Split(" "c)
    Dim justifiedText As String = ""
    Dim line As String = ""

    For Each word As String In words
        If (line + word).Length > maxCharsPerLine Then
            Dim faltantes As Integer = maxCharsPerLine - line.Trim().Length
            Dim subwords As String() = line.Trim().Split(" ")
            Dim limit = subwords.Length - 1
            Dim aux = 0
            For i As Integer = 0 To faltantes - 1
                If aux < limit Then
                    subwords(aux) += "&nbsp;"
                    aux += 1
                Else
                    aux = 0
                    subwords(aux) += "&nbsp;"
                End If
            Next
            For Each s As String In subwords
                justifiedText += s + " "
            Next
            justifiedText += "<br />"
            line = ""
        End If

        line += word + " "
    Next

    justifiedText += line.Trim()
    Return justifiedText
End Function
```

 - text (String): Este es el bloque de texto que deseas justificar y formatear en HTML. Es el texto que quieres que la función divida en líneas y ajuste según el límite de caracteres por línea.

 - maxCharsPerLine (Integer): Este es el número máximo de caracteres permitidos por línea. La función intentará justificar el texto para que cada línea tenga aproximadamente esta cantidad de caracteres, agregando espacios no rompibles (&nbsp;) según sea necesario.

Ejemplo de uso :
```vb
=Code.TextOnlyJustifyHTML(
"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Maecenas ac lorem nec diam auctor aliquam sit amet quis augue. Praesent pharetra dolor malesuada eros laoreet volutpat. Nulla ac rhoncus lorem. Pellentesque vel tempus ex, ut sagittis eros. Curabitur rutrum lacus risus, vel molestie augue euismod quis. Nunc dapibus convallis lectus, eu mattis nisi dictum eget. Morbi quis enim eget dolor ultricies ultrices id in lacus. Sed nec tellus ut enim lacinia malesuada. Etiam sit amet sollicitudin erat." +
"Curabitur auctor hendrerit pulvina. Ut imperdiet arcu vitae rutrum iaculis. Donec at est cursus, laoreet turpis quis, eleifend odio. Donec at mattis magna. Praesent odio ligula, tincidunt et erat non, laoreet semper metus. Praesent hendrerit elit quis commodo commodo. Pellentesque non porttitor ipsum. Aenean eu mauris ut lorem varius tincidunt sed id dui. Proin euismod nec orci quis rutrum. Nunc gravida gravida posuere. Suspendisse nec porttitor orci. Donec consectetur nisl eu tempor feugiat. Cras vulputate sem quis mauris mattis, nec tincidunt orci pulvinar. Proin viverra ipsum tellus, quis accumsan magna fringilla sit amet. Aliquam egestas augue nec lacus suscipit sagittis. Nulla quis sodales lorem."
, 155)
```

### Función TextOnlyJustifyHtmlBold

Esta función combina las características de justificación de texto y formato en negrita. Realiza la justificación del texto y resalta palabras específicas en negrita. El resultado se devuelve en formato HTML.

```vb
Function TextOnlyJustifyHtmlBold(text As String, max_chars_per_line As Integer) As String
    Dim words As String() = text.Split()
    Dim justified_text As String = ""
    Dim line As String = ""

    For Each word As String In words
        If line.Length + word.Length > max_chars_per_line Then
            Dim faltantes As Integer = max_chars_per_line - line.Trim().Length
            Dim subwords As String() = line.Trim().Split()
            Dim limit As Integer = subwords.Length - 1
            Dim aux As Integer = 0

            For i As Integer = 1 To faltantes
                If aux < limit Then
                    subwords(aux) += "&nbsp;"
                    aux += 1
                Else
                    aux = 0
                    subwords(aux) += "&nbsp;"
                End If
            Next

            For Each s As String In subwords
                justified_text += s & " "
            Next
            justified_text += "<br />"
            line = ""

        End If

        line += word & " "
    Next

    justified_text += line.Trim()

    Dim bold_words As String() = {"Lorem", "dolor"}
    justified_text = ApplyBoldFormat(justified_text, bold_words)

    Return justified_text
End Function
```

 - ApplyBoldFormat: Esta función recorre el texto y reemplaza las palabras específicas con su versión en negrita. Esto es útil para resaltar ciertas palabras clave en el contenido.

 - TextOnlyJustifyHtmlBold: Esta función toma un texto y un número máximo de caracteres por línea como entrada. Procesa el texto palabra por palabra, justificándolo para que se ajuste al ancho deseado. Si una palabra no cabe en la línea actual, se agregan espacios adicionales para ajustar el texto. Luego se agrega un salto de línea y se procede a la siguiente línea.

Ejemplo de uso :
```vb
=Code.TextOnlyJustifyHtmlBold(
"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Maecenas ac lorem nec diam auctor aliquam sit amet quis augue. Praesent pharetra dolor malesuada eros laoreet volutpat. Nulla ac rhoncus lorem. Pellentesque vel tempus ex, ut sagittis eros. Curabitur rutrum lacus risus, vel molestie augue euismod quis. Nunc dapibus convallis lectus, eu mattis nisi dictum eget. Morbi quis enim eget dolor ultricies ultrices id in lacus. Sed nec tellus ut enim lacinia malesuada. Etiam sit amet sollicitudin erat." +
"Curabitur auctor hendrerit pulvina. Ut imperdiet arcu vitae rutrum iaculis. Donec at est cursus, laoreet turpis quis, eleifend odio. Donec at mattis magna. Praesent odio ligula, tincidunt et erat non, laoreet semper metus. Praesent hendrerit elit quis commodo commodo. Pellentesque non porttitor ipsum. Aenean eu mauris ut lorem varius tincidunt sed id dui. Proin euismod nec orci quis rutrum. Nunc gravida gravida posuere. Suspendisse nec porttitor orci. Donec consectetur nisl eu tempor feugiat. Cras vulputate sem quis mauris mattis, nec tincidunt orci pulvinar. Proin viverra ipsum tellus, quis accumsan magna fringilla sit amet. Aliquam egestas augue nec lacus suscipit sagittis. Nulla quis sodales lorem."
, 155)
```

### Función ApplyBoldFormat

Esta función toma una cadena de texto y una serie de palabras a resaltar. Devuelve la cadena original con las palabras resaltadas en negrita utilizando etiquetas `<b>` de HTML.

```vb
Function ApplyBoldFormat(text As String, ParamArray bold_words() As String) As String
    For Each bold_word As String In bold_words
        text = text.Replace(bold_word, "<b>" & bold_word & "</b>")
    Next
    Return text
End Function

```

## Cómo Utilizar las Funciones

1. Copia las funciones proporcionadas en tu informe de SSRS, preferiblemente en el módulo de código del informe.
2. Llama a la función `TextOnlyJustify`, `TextOnlyJustifyHTML` o `TextOnlyJustifyHtmlBold`, pasando los argumentos necesarios como texto, fuente, ancho máximo, etc.
3. Utiliza la cadena devuelta en tu informe de SSRS.

Recuerda que estas funciones son específicas para justificar texto con fuentes monoespaciadas y se deben usar en contextos donde este tipo de justificación sea necesario.

Esperamos que estas funciones te sean útiles para lograr la presentación deseada en tus informes de SSRS. ¡No dudes en ajustarlas según tus necesidades específicas!
