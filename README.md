# go-excel-util

A very light go binary that will replaces the content of a cell in an excel file

Syntax ``go-excel-util <file-name> <sheet#> <row#> <col#> <replacement string>``

## Sample Usage

### With an excel file named test.xlsx with the following contents

| Sheet 1 | | |
|-------------|-------------|-------------|
| index 1,1,1 | index 1,2,1 | index 1,3,1 |
| index 1,1,2 | index 1,2,2 | index 1,3,2 |
| index 1,1,3 | index 1,2,3 | index 1,3,3 |

| Sheet 2 | | |
|-------------|-------------|-------------|
| index 2,1,1 | index 2,2,1 | index 2,3,1 |
| index 2,1,2 | index 2,2,2 | index 2,3,2 |
| index 2,1,3 | index 2,2,3 | index 2,3,3 |

### Running the following


```bash
$ go-excel-util test.xlsx 1 1 1 Replaced!
2019/07/25 09:31:20 Repling 'index 1,1,1' with 'Replaced!'

$ go-excel-util test.xlsx 1 1 1 'hatdog'
2019/07/25 09:31:58 Replacing 'Replaced!' with 'hatdog'

$ go-excel-util test.xlsx 2 2 2 'hatdog sa sheet2'
2019/07/25 09:32:57 Repling 'index 2,2,2' with 'hatdog sa sheet2'

$ go-excel-util test.xlsx 2 3 2 'isa pang hatdog sa sheet2!'
2019/07/25 09:33:49 Repling 'index 2,3,2' with 'isa pang hatdog sa sheet2!'

```

### Will result into

| Sheet 1 | | |
|-------------|-------------|-------------|
| hatdog | index 1,2,1 | index 1,3,1 |
| index 1,1,2 | index 1,2,2 | index 1,3,2 |
| index 1,1,3 | index 1,2,3 | index 1,3,3 |

| Sheet 2 | | |
|-------------|-------------|-------------|
| index 2,1,1 | index 2,2,1 | index 2,3,1 |
| index 2,1,2 | hatdog sa sheet2 | isa pang hatdog sa sheet2! |
| index 2,1,3 | index 2,2,3 | index 2,3,3 |
