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

Add the dependency (for all available versions see [https://jitpack.io/#Serbroda/poi-data-template](https://jitpack.io/#Serbroda/thymeleaf-component-dialect)).

```html
<dependency>
    <groupId>com.github.Serbroda</groupId>
    <artifactId>poi-data-template</artifactId>
    <version>VERSION</version>
</dependency>
```

Usage
------

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
        d.setId((int) row.getCell(0).getNumericCellValue());
        d.setFirstName(row.getCell(1).getStringCellValue());
        d.setLastName(row.getCell(2).getStringCellValue());
        d.setDate(row.getCell(3).getDateCellValue());
        d.setSales(
                new BigDecimal(row.getCell(4).getNumericCellValue()));
        return d;
    }, true);

```
