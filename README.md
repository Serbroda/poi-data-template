POI data template
===========================

[![Build Status](https://travis-ci.org/Serbroda/poi-data-template.svg?branch=develop)](https://travis-ci.org/Serbroda/poi-data-template)
[![jitpack](https://jitpack.io/v/Serbroda/poi-data-template.svg)](https://jitpack.io/#Serbroda/poi-data-template)

Installation
------

Add the jitpack repository.

```html
<repositories>
    <repository>
        <id>jitpack.io</id>
        <url>https://jitpack.io</url>
    </repository>
</repositories>
```

Add the dependency (for all available versions see [https://jitpack.io/#Serbroda/poi-data-template](https://jitpack.io/#Serbroda/poi-data-template)).

```html
<dependency>
    <groupId>com.github.Serbroda</groupId>
    <artifactId>poi-data-template</artifactId>
    <version>VERSION</version>
</dependency>
```

Usage
------

test.xlsx
```text
+----+------------+------------+------------+---------+
| ID | FIRST_NAME | LAST_NAME  |    DATE    |  SALES  |
+----+------------+------------+------------+---------+
|  1 | Max        | Mustermann | 01.02.2018 | 12,95 € |
|  2 | John       | Doe        | 27.06.2018 | 29,50 € |
+----+------------+------------+------------+---------+

```

```java
 public static class Data {
    public int id;
    public String firstName;
    public String lastName;
    public Date date;
    public BigDecimal sales;
}

private ExcelSource source = new ExcelFileSource("test.xlsx");

List<Data> data = new ExcelDataTemplate()
    .read(source, 0, row -> {
        Data d = new Data();
        d.id = (int) row.getCell(0).getNumericCellValue();
        d.firstName = row.getCell(1).getStringCellValue();
        d.lastName = row.getCell(2).getStringCellValue();
        d.date = row.getCell(3).getDateCellValue();
        d.sales = new BigDecimal(row.getCell(4).getNumericCellValue());
        return d;
    }, true);

```
