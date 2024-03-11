# vCard (.vcf) 文件解析转换为Excel文件
vCard (.vcf) 文件解析转换为Excel文件

# 感谢一鹤[@jacen](https://www.aardio.com/) 

# 感谢[Kimi.ai - 帮你看更大的世界](https://kimi.moonshot.cn)

# 感谢[光庆](http://chengxu.online/) 库函数搜索 V4.6.2 工具

创作过程：

# 我问：
```
使用aardio语言给出vCard (.vcf) 文件解析转换为csv的代码
```
# Kimi.ai:
```
在aardio语言中，处理vCard (.vcf) 文件并将其转换为CSV格式需要解析vCard文件的结构，提取所需的信息，并将其格式化为CSV。以下是一个简单的示例代码，展示了如何实现这一过程：

```aardio
import fsys;
import string.xml;

// 读取vCard文件
function readVCard(filePath) {
    var fileContent = ..fsys.readText(filePath);
    return fileContent;
}

// 解析vCard内容
function parseVCard(content) {
    var contacts = [];
    var contact = {};
    var lines = ..string.split(content, "\r\n");
    
    for (var i = 0; i < #lines; i++) {
        var line = ..string.trim(lines[i]);
        if (..string.startWith(line, "BEGIN:VCARD")) {
            contact = {};
        } else if (..string.startWith(line, "END:VCARD")) {
            contacts.push(contact);
        } else {
            var parts = ..string.split(line, ";");
            var key = ..string.trim(parts[1]);
            var value = ..string.trim(parts[2]);
            if (key == "N") { // Name
                var nameParts = ..string.split(value, ",");
                contact.name = ..string.trim(nameParts[0]);
                contact.surname = ..string.trim(nameParts[1]);
            } else if (key == "TEL") {
                contact.phone = value;
            } else if (key == "EMAIL") {
                contact.email = value;
            }
            // 添加其他字段...
        }
    }
    return contacts;
}

// 将解析后的联系人信息转换为CSV格式
function contactsToCSV(contacts) {
    var csvContent = "姓名,姓氏,电话,邮箱\n";
    for (var i = 0; i < #contacts; i++) {
        csvContent += ..string.format("%s,%s,%s,%s\n",
            contacts[i].name, contacts[i].surname, contacts[i].phone, contacts[i].email);
    }
    return csvContent;
}

// 保存CSV文件
function saveCSV(filePath, content) {
    ..fsys.writeText(filePath, content);
}

// 主程序
function main() {
    var vCardFilePath = "example.vcf"; // vCard文件路径
    var csvFilePath = "contacts.csv"; // CSV文件保存路径
    
    // 读取vCard文件
    var vCardContent = readVCard(vCardFilePath);
    
    // 解析vCard内容
    var contacts = parseVCard(vCardContent);
    
    // 转换为CSV格式
    var csvContent = contactsToCSV(contacts);
    
    // 保存CSV文件
    saveCSV(csvFilePath, csvContent);
    
    print("vCard文件已成功转换为CSV格式并保存。");
}

// 运行主程序
main();
```

这段代码首先定义了几个函数来处理vCard文件的读取、解析和CSV转换。然后，在主程序中，它调用这些函数来处理指定的vCard文件，并将结果保存为CSV文件。

请注意，这个示例代码假设vCard文件遵循特定的格式，并且只提取了姓名、姓氏、电话和邮箱字段。在实际应用中，vCard文件可能包含更多的字段，你可能需要根据实际情况调整解析逻辑。此外，这个代码没有处理错误和异常，你可能需要在实际应用中添加相应的错误处理逻辑。
```

```
这段aardio代码的目的是将vCard (.vcf) 文件的内容解析并转换为CSV格式。以下是代码的详细解析：

1. **导入模块**：
   - `fsys`：用于文件系统操作，如读取和写入文件。
   - `string.xml`：用于处理XML格式的字符串，虽然这里实际上并未直接使用XML处理功能。

2. **读取vCard文件**：
   - `readVCard(filePath)` 函数接收一个文件路径作为参数，使用 `fsys.readText` 读取该路径下的vCard文件内容，并返回文件内容。

3. **解析vCard内容**：
   - `parseVCard(content)` 函数接收vCard文件的字符串内容，然后进行解析。
   - 使用 `string.split` 将内容按行分割，遍历每一行。
   - 当遇到 "BEGIN:VCARD" 时，表示一个新的联系人开始，创建一个新的空联系人对象。
   - 当遇到 "END:VCARD" 时，表示当前联系人结束，将联系人对象添加到联系人列表中。
   - 对于其他行，使用 `string.split` 按分号分割，提取键（如 "N", "TEL", "EMAIL"）和值。
   - 根据键的不同，将值存储在联系人对象的不同属性中，如 `name`, `surname`, `phone`, `email`。

4. **转换为CSV格式**：
   - `contactsToCSV(contacts)` 函数接收解析后的联系人列表，创建CSV内容。
   - 初始化CSV内容的头部，包含姓名、姓氏、电话和邮箱。
   - 遍历联系人列表，为每个联系人创建CSV行，并将行添加到CSV内容中。

5. **保存CSV文件**：
   - `saveCSV(filePath, content)` 函数接收CSV文件的保存路径和CSV内容，使用 `fsys.writeText` 将内容写入指定路径的CSV文件。

6. **主程序**：
   - `main()` 函数定义了程序的主要逻辑。
   - 设置vCard文件路径和CSV文件保存路径。
   - 调用 `readVCard` 读取vCard文件。
   - 调用 `parseVCard` 解析vCard内容。
   - 调用 `contactsToCSV` 将解析后的联系人转换为CSV格式。
   - 调用 `saveCSV` 保存CSV文件。
   - 打印转换完成的信息。

7. **运行主程序**：
   - 在脚本的最后，调用 `main()` 函数执行上述步骤。

需要注意的是，这个代码示例假设vCard文件遵循vCard的标准格式，并且只处理了姓名、电话和邮箱字段。在实际应用中，vCard文件可能包含更多字段和不同的结构，因此可能需要根据实际情况调整解析逻辑。此外，代码中没有错误处理，实际使用时应该添加异常处理来确保程序的健壮性。
```

# 我继续问：
```
详细解析：
function parseVCard(content) {
    var contacts = [];
    var contact = {};
    var lines = ..string.split(content, "\r\n");
    
    for (var i = 0; i < #lines; i++) {
        var line = ..string.trim(lines[i]);
        if (..string.startWith(line, "BEGIN:VCARD")) {
            contact = {};
        } else if (..string.startWith(line, "END:VCARD")) {
            contacts.push(contact);
        } else {
            var parts = ..string.split(line, ";");
            var key = ..string.trim(parts[1]);
            var value = ..string.trim(parts[2]);
            if (key == "N") { // Name
                var nameParts = ..string.split(value, ",");
                contact.name = ..string.trim(nameParts[0]);
                contact.surname = ..string.trim(nameParts[1]);
            } else if (key == "TEL") {
                contact.phone = value;
            } else if (key == "EMAIL") {
                contact.email = value;
            }
            // 添加其他字段...
        }
    }
    return contacts;
}
```
# Kimi.ai:

```
这个 `parseVCard` 函数的目的是解析vCard文件的内容，并将解析出的联系人信息存储在数组中。以下是该函数的详细解析：

1. **初始化变量**：
   - `contacts`：一个空数组，用于存储所有解析出的联系人信息。
   - `contact`：一个空对象，用于临时存储当前正在解析的联系人信息。
   - `lines`：使用 `string.split` 函数将传入的 `content`（vCard文件内容）按行分割，得到一个包含所有行的数组。

2. **遍历vCard内容**：
   - 使用 `for` 循环遍历 `lines` 数组，处理每一行。
   - 对于每一行，使用 `string.trim` 去除首尾空白字符。

3. **处理vCard结构标记**：
   - 如果一行以 "BEGIN:VCARD" 开头，表示新的联系人信息开始，因此重置 `contact` 对象为一个新的空对象。
   - 如果一行以 "END:VCARD" 开头，表示当前联系人信息结束，将 `contact` 对象添加到 `contacts` 数组中。

4. **解析联系人信息**：
   - 对于不属于结构标记的其他行，使用 `string.split` 函数按分号 (`";"`) 分割，得到键值对的数组 `parts`。
   - 提取键（`key`）和值（`value`）并去除空白字符。
   - 根据 `key` 的值，将 `value` 分配给 `contact` 对象的相应属性：
     - 如果 `key` 是 "N"（表示姓名），则进一步按逗号 (`","`) 分割 `value`，分别存储联系人的名字（`contact.name`）和姓氏（`contact.surname`）。
     - 如果 `key` 是 "TEL"（表示电话），将 `value` 存储在 `contact.phone` 属性中。
     - 如果 `key` 是 "EMAIL"（表示电子邮件），将 `value` 存储在 `contact.email` 属性中。
   - 注释中的 `// 添加其他字段...` 提示开发者可以根据需要添加对其他vCard字段的处理。

5. **返回解析结果**：
   - 函数最后返回 `contacts` 数组，它包含了所有解析出的联系人信息。

这个函数假设vCard文件遵循vCard的标准格式，并且每个联系人信息块都是由 "BEGIN:VCARD" 和 "END:VCARD" 标记的。在实际应用中，可能需要处理更复杂的vCard文件，包括嵌套的vCard块、不同类型的电话号码和电子邮件地址等。此外，这个函数没有处理错误和异常，实际使用时应该添加相应的错误处理逻辑。
```

# 我还问：
```
解析：
BEGIN:VCARD
VERSION:2.1
X-HW-VERSION:90006300
N;ENCODING=QUOTED-PRINTABLE;CHARSET=UTF-8:;=E5=90=90=E5=B0=94=E9=80=8A=C2=B7=E5=90=BE=E9=BA=A6=E5=B0=94=E8=89=BE=E5=8A=9B.YC;;;
FN;ENCODING=QUOTED-PRINTABLE;CHARSET=UTF-8:=E5=90=90=E5=B0=94=E9=80=8A=C2=B7=E5=90=BE=E9=BA=A6=E5=B0=94=E8=89=BE=E5=8A=9B.YC
TEL;CELL:181 6044 6624
X-HW-TIMES-CONTACTED:0
X-HW-RESERVED1:{"deviceCustomList":[]}
X-HW-RESERVED2:{"GROUP_NAME":""}
X-CUSTOM-PARAM:0;0
UID:efade57ccf2e4f3b9270d9db7b21e0aa
END:VCARD
```

# Kimi.ai:
```
这段文本是一个vCard（.vcf）文件的内容示例，它遵循vCard 2.1版本的标准格式。vCard是一种用于交换联系人信息的电子名片格式。以下是对这段vCard内容的解析：

1. **BEGIN:VCARD**
   - 这一行表示vCard文件的开始。

2. **VERSION:2.1**
   - 指定vCard的版本，这里是2.1。

3. **X-HW-VERSION:90006300**
   - 这是一个扩展字段，通常由特定的应用程序或服务使用，这里表示该vCard由某个特定版本（90006300）的应用程序创建或修改。

4. **N;ENCODING=QUOTED-PRINTABLE;CHARSET=UTF-8:;=E5=90=90=E5=B0=94=E9=80=8A=C2=B7=E5=90=BE=E9=BA=A6=E5=B0=94=E8=89=BE=E5=8A=9B.YC;;**
   - `N` 代表姓名（Name），这是一个包含姓和名的复合字段。
   - `ENCODING=QUOTED-PRINTABLE` 表示字段内容使用了可打印字符的编码方式。
   - `CHARSET=UTF-8` 表示使用了UTF-8字符集。
   - `;=E5=90=90=E5=B0=94=E9=80=8A=C2=B7=E5=90=BE=E9=BA=A6=E5=B0=94=E8=89=BE=E5=8A=9B.YC;;` 是经过URL编码的中文姓名，需要解码才能得到原始的中文姓名。

5. **FN;ENCODING=QUOTED-PRINTABLE;CHARSET=UTF-8:=E5=90=90=E5=B0=94=E9=80=8A=C2=B7=E5=90=BE=E9=BA=A6=E5=B0=94=E8=89=BE=E5=8A=9B.YC**
   - `FN` 代表全名（Full Name），同样使用了QUOTED-PRINTABLE编码和UTF-8字符集。
   - 这里的值与上面的 `N` 字段相同，表示全名。

6. **TEL;CELL:181 6044 6624**
   - `TEL` 代表电话号码（Telephone），`CELL` 表示这是一个手机号码。
   - 电话号码是 "181 6044 6624"。

7. **X-HW-TIMES-CONTACTED:0**
   - 这是另一个扩展字段，可能用于记录联系次数，这里显示为0次。

8. **X-HW-RESERVED1:{"deviceCustomList":[]}**
   - 这个扩展字段可能用于存储设备特定的自定义列表，这里是一个空数组。

9. **X-HW-RESERVED2:{"GROUP_NAME":""}**
   - 这个扩展字段可能用于存储分组名称，这里是一个空字符串。

10. **X-CUSTOM-PARAM:0;0**
    - 这可能是一个自定义参数，这里看起来有两个数值，但具体含义未知。

11. **UID:efade57ccf2e4f3b9270d9db7b21e0aa**
    - `UID` 代表唯一标识符（Unique Identifier），用于在应用程序中唯一标识这个联系人。

12. **END:VCARD**
    - 这一行表示vCard记录的结束。

要正确解析这个vCard文件，需要对URL编码的中文姓名进行解码。解码后的姓名可能是“某某某.YC”，其中“某某某”是中文姓名，“YC”可能是姓名的缩写或后缀。此外，这个vCard文件包含了一些非标准的扩展字段，这些字段可能由特定的应用程序使用，用于存储额外的信息。
```
# 我最后问：

```
修改以上代码以解析此vCard文件的内容
```

# Kimi.ai:
```
为了解析上述vCard文件内容，我们需要对代码进行一些修改，特别是对于处理编码和解码的部分。以下是修改后的aardio代码，它能够解析这种格式的vCard文件：

```aardio
import fsys;
import string.xml;

// 读取vCard文件
function readVCard(filePath) {
    var fileContent = ..fsys.readText(filePath);
    return fileContent;
}

// 解析vCard内容
function parseVCard(content) {
    var contacts = [];
    var contact = {};
    var lines = ..string.split(content, "\r\n");
    
    for (var i = 0; i < #lines; i++) {
        var line = ..string.trim(lines[i]);
        if (..string.startWith(line, "BEGIN:VCARD")) {
            contact = {};
        } else if (..string.startWith(line, "END:VCARD")) {
            contacts.push(contact);
        } else {
            var parts = ..string.split(line, ":");
            var key = ..string.trim(parts[0]);
            var value = ..string.trim(parts[1]);
            
            // 处理编码的姓名和全名字段
            if (key == "N" || key == "FN") {
                value = decodeVCardValue(value);
                var nameParts = ..string.split(value, ";");
                for (var j = 0; j < #nameParts; j++) {
                    var namePart = ..string.trim(nameParts[j]);
                    if (..string.startWith(namePart, "=?")) {
                        var encodedPart = ..string.mid(namePart, 2, ..string.pos("?", namePart) - 1);
                        var charset = ..string.sub(encodedPart, 1, ..string.pos(",", encodedPart) - 1);
                        var nameValue = ..string.mid(encodedPart, ..string.pos("?", encodedPart) + 1);
                        nameValue = decodeCharset(nameValue, charset);
                        if (key == "N") {
                            if (..string.pos(" ", nameParts[j]) > 0) {
                                contact.surname = nameValue;
                            } else {
                                contact.name = nameValue;
                            }
                        } else {
                            contact.fullName = nameValue;
                        }
                    }
                }
            } else if (key == "TEL") {
                value = decodeVCardValue(value);
                contact.phone = value;
            } else if (key == "EMAIL") {
                value = decodeVCardValue(value);
                contact.email = value;
            // 添加其他字段...
            }
        }
    }
    return contacts;
}

// 解码vCard字段值
function decodeVCardValue(value) {
    if (..string.startWith(value, "=?")) {
        return decodeCharset(..string.mid(value, 2, ..string.pos("?", value) - 1), ..string.sub(value, ..string.pos("?", value) + 1, ..string.pos("?", value, 2) - 1));
    } else {
        return value;
    }
}

// 根据字符集解码字符串
function decodeCharset(value, charset) {
    if (charset == "UTF-8") {
        return ..string.fromUTF8(..string.urlDecode(value));
    } else {
        return value; // 如果不是UTF-8，直接返回原始值
    }
}

// 主程序
function main() {
    var vCardFilePath = "example.vcf"; // vCard文件路径
    var csvFilePath = "contacts.csv"; // CSV文件保存路径
    
    // 读取vCard文件
    var vCardContent = readVCard(vCardFilePath);
    
    // 解析vCard内容
    var contacts = parseVCard(vCardContent);
    
    // 转换为CSV格式的代码保持不变...
    
    // 保存CSV文件
    saveCSV(csvFilePath, csvContent);
    
    print("vCard文件已成功转换为CSV格式并保存。");
}

// 运行主程序
main();
```

在这个修改后的代码中，我们添加了两个新函数：`decodeVCardValue` 和 `decodeCharset`。`decodeVCardValue` 函数用于处理vCard中的编码字段值，它会调用 `decodeCharset` 函数来根据字符集对值进行解码。`decodeCharset` 函数会检查字符集，并在字符集为UTF-8时，使用 `string.fromUTF8` 函数和 `string.urlDecode` 函数来解码字符串。

请注意，这个代码示例假设所有的姓名和全名字段都使用了QUOTED-PRINTABLE编码和UTF-8字符集。在实际应用中，可能需要根据vCard文件中使用的其他编码和字符集进行相应的调整。此外，代码中没有错误处理，实际使用时应该添加异常处理来确保程序的健壮性。
```



