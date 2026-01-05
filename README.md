# Leitor de Arquivos Excel

Uma aplicação Java para ler arquivos Excel (.xlsx e .xls) e converter as linhas em objetos Java.

## 📁 Estrutura do Projeto

O projeto segue uma arquitetura limpa com pacotes bem definidos:

```
src/main/java/org/example/
├── Main.java                    # Ponto de entrada da aplicação
├── model/
│   └── Person.java              # Classe modelo de dados
├── reader/
│   ├── RowMapper.java           # Interface para mapeamento de linhas
│   └── ExcelReader.java         # Classe que lê arquivos Excel
├── mapper/
│   └── PersonRowMapper.java     # Mapper específico para Person
└── util/
    └── CellValueExtractor.java  # Utilitário para extrair valores das células
```

## 📦 Dependências

Este projeto usa **Apache POI** para manipulação de arquivos Excel:

```xml
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>5.2.5</version>
</dependency>

<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.2.5</version>
</dependency>
```

- `poi` - Biblioteca principal para documentos Office
- `poi-ooxml` - Suporte para arquivos .xlsx (Excel 2007+)

## 🚀 Como Executar

1. Certifique-se de ter o **Java 17+** instalado
2. Abra o projeto no **IntelliJ IDEA** (ele baixará automaticamente as dependências do Maven)
3. Execute a classe `Main`
4. A aplicação lerá o arquivo `docs/exemplo.xlsx` e exibirá os dados no console

## 🔧 Como Adaptar para Diferentes Arquivos Excel

A estrutura do código foi projetada para ser facilmente adaptável a qualquer arquivo Excel. Siga os passos abaixo:

### Passo 1: Criar uma Classe Modelo

Crie uma nova classe no pacote `org.example.model` que represente os dados do seu Excel.

**Exemplo:** Se você tem um Excel com produtos (Código, Nome, Preço, Quantidade):

```java
package org.example.model;

/**
 * Classe modelo representando um produto.
 */
public class Product {
    
    private String code;      // Coluna A - Código do produto
    private String name;      // Coluna B - Nome do produto
    private Double price;     // Coluna C - Preço
    private Integer quantity; // Coluna D - Quantidade em estoque
    
    // Construtor padrão
    public Product() {
    }
    
    // Construtor com parâmetros
    public Product(String code, String name, Double price, Integer quantity) {
        this.code = code;
        this.name = name;
        this.price = price;
        this.quantity = quantity;
    }
    
    // Getters e Setters
    public String getCode() { return code; }
    public void setCode(String code) { this.code = code; }
    
    public String getName() { return name; }
    public void setName(String name) { this.name = name; }
    
    public Double getPrice() { return price; }
    public void setPrice(Double price) { this.price = price; }
    
    public Integer getQuantity() { return quantity; }
    public void setQuantity(Integer quantity) { this.quantity = quantity; }
    
    @Override
    public String toString() {
        return "Product{code='" + code + "', name='" + name + 
               "', price=" + price + ", quantity=" + quantity + "}";
    }
}
```

### Passo 2: Criar um RowMapper

Crie um novo mapper no pacote `org.example.mapper` que implemente a interface `RowMapper<T>`.

O mapper é responsável por:
- Definir de qual coluna cada dado será extraído
- Converter os valores das células para os tipos corretos

```java
package org.example.mapper;

import org.apache.poi.ss.usermodel.Row;
import org.example.model.Product;
import org.example.reader.RowMapper;
import org.example.util.CellValueExtractor;

/**
 * Mapper para converter linhas do Excel em objetos Product.
 */
public class ProductRowMapper implements RowMapper<Product> {
    
    // Defina os índices das colunas (0 = A, 1 = B, 2 = C, etc.)
    private static final int CODE_COLUMN = 0;     // Coluna A
    private static final int NAME_COLUMN = 1;     // Coluna B
    private static final int PRICE_COLUMN = 2;    // Coluna C
    private static final int QUANTITY_COLUMN = 3; // Coluna D
    
    @Override
    public Product mapRow(Row row) {
        Product product = new Product();
        
        // Extrair cada valor usando o CellValueExtractor
        product.setCode(CellValueExtractor.getStringValue(row.getCell(CODE_COLUMN)));
        product.setName(CellValueExtractor.getStringValue(row.getCell(NAME_COLUMN)));
        product.setPrice(CellValueExtractor.getDoubleValue(row.getCell(PRICE_COLUMN)));
        product.setQuantity(CellValueExtractor.getIntegerValue(row.getCell(QUANTITY_COLUMN)));
        
        return product;
    }
}
```

### Passo 3: Usar o Leitor

Na sua classe Main (ou onde precisar), use o `ExcelReader` com o seu mapper:

```java
// Criar o leitor e o mapper
ExcelReader reader = new ExcelReader();
ProductRowMapper mapper = new ProductRowMapper();

// Ler o arquivo Excel
List<Product> products = reader.readFile("caminho/para/produtos.xlsx", mapper);

// Usar os dados
for (Product product : products) {
    System.out.println(product);
}
```

## 📊 Extração de Valores das Células

A classe `CellValueExtractor` fornece métodos para extrair diferentes tipos de dados:

| Método | Tipo de Retorno | Descrição | Quando Usar |
|--------|-----------------|-----------|-------------|
| `getStringValue(Cell)` | String | Extrai valores de texto | Nomes, códigos, descrições |
| `getIntegerValue(Cell)` | Integer | Extrai números inteiros | Quantidades, idades, IDs |
| `getDoubleValue(Cell)` | Double | Extrai números decimais | Preços, percentuais |
| `getBooleanValue(Cell)` | Boolean | Extrai valores verdadeiro/falso | Status, flags |
| `getDateValue(Cell)` | Date | Extrai valores de data | Datas de nascimento, vencimentos |

**Importante:** Todos os métodos tratam células nulas e diferentes tipos de células de forma segura, retornando `null` quando o valor não pode ser extraído.

## 📋 Exemplo Completo: Lendo um Excel de Funcionários

Imagine que você tem um Excel com a seguinte estrutura:

| Nome | Cargo | Salário | Data de Admissão | Ativo |
|------|-------|---------|------------------|-------|
| João Silva | Desenvolvedor | 5500.00 | 15/03/2020 | Sim |
| Maria Santos | Analista | 4800.50 | 22/07/2021 | Sim |

**1. Criar o modelo Employee.java:**

```java
package org.example.model;

import java.util.Date;

public class Employee {
    private String name;
    private String position;
    private Double salary;
    private Date hireDate;
    private Boolean active;
    
    // Getters, setters e toString...
}
```

**2. Criar o EmployeeRowMapper.java:**

```java
package org.example.mapper;

import org.apache.poi.ss.usermodel.Row;
import org.example.model.Employee;
import org.example.reader.RowMapper;
import org.example.util.CellValueExtractor;

public class EmployeeRowMapper implements RowMapper<Employee> {
    
    @Override
    public Employee mapRow(Row row) {
        Employee emp = new Employee();
        emp.setName(CellValueExtractor.getStringValue(row.getCell(0)));
        emp.setPosition(CellValueExtractor.getStringValue(row.getCell(1)));
        emp.setSalary(CellValueExtractor.getDoubleValue(row.getCell(2)));
        emp.setHireDate(CellValueExtractor.getDateValue(row.getCell(3)));
        emp.setActive(CellValueExtractor.getBooleanValue(row.getCell(4)));
        return emp;
    }
}
```

**3. Usar no Main:**

```java
ExcelReader reader = new ExcelReader();
List<Employee> employees = reader.readFile("funcionarios.xlsx", new EmployeeRowMapper());
```

## ⚙️ Opções Avançadas do ExcelReader

### Ler uma planilha específica

```java
// Ler a segunda planilha (índice 1)
List<Person> people = reader.readFile("arquivo.xlsx", mapper, 1, 1);
// Parâmetros: arquivo, mapper, linhas de cabeçalho, índice da planilha
```

### Arquivo sem cabeçalho

```java
// Se o Excel não tem linha de cabeçalho
List<Person> people = reader.readFile("arquivo.xlsx", mapper, 0);
```

### Listar planilhas disponíveis

```java
List<String> sheets = reader.getSheetNames("arquivo.xlsx");
System.out.println("Planilhas: " + sheets);
```

## 📄 Formatos de Excel Suportados

| Formato | Extensão | Versão do Excel |
|---------|----------|-----------------|
| XLSX | .xlsx | Excel 2007 e superior (recomendado) |
| XLS | .xls | Excel 97-2003 (legado) |

## ✨ Características do Projeto

- ✅ Mapeamento genérico de linhas usando padrão Strategy
- ✅ Suporte a múltiplas planilhas
- ✅ Pular linhas de cabeçalho configurável
- ✅ Detecção e pulo de linhas vazias
- ✅ Tratamento de erros abrangente
- ✅ Código amplamente documentado

## 🔍 Dicas de Uso

1. **Índices de colunas começam em 0**: Coluna A = 0, B = 1, C = 2, etc.

2. **Crie constantes para os índices**: Facilita a manutenção se a ordem das colunas mudar.

3. **Use o tipo correto**: Se uma coluna pode ter decimais, use `Double`. Se sempre será inteiro, use `Integer`.

4. **Trate valores nulos**: Os métodos do `CellValueExtractor` retornam `null` para células vazias.
