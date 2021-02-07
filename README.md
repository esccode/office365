# office365
w polskim exel separotorem jest srednik
w angielskim exel separatorem jest przecinek

## Exel
https://exceljet.net/excel-functions/excel-lookup-function

### Data Validation
data -> data Validation

### Excel lookup(wyszukaj) v vlookup(wszukaj pionowo) Functions
=LOOKUP (lookup_value, lookup_vector, [result_vector])
=LOOKUP(G7,C6:C10,D6:D10)
=LOOKUP(G7,C6:C10)

=VLOOKUP (value, table, col_index, [range_lookup])
=VLOOKUP(D4,A10:E11,4,FALSE)

=VLOOKUP(HostList[@Hostname],Sheet1!D:E,2,FALSE)

### Odwołania względne i bezwględne
=$B$1*$A$2
blokujemy kolumne lub wiersze lub obje jednoczesnie

### =jezeli moze byz z IFERROR
=jazeli(B2>=A4;"jest wieksze";"nie jest wieksze")
=IF(logical_test, [value_if_true], [value_if_false])

### =warunki (IFS) z IFERROR
=IFERROR(IFS(A10=23,"wynik pierwszego warunku",A12>=25,"wynik drugiego warunku"),"jezeli nie spelniony zaden warunek")
=IFERROR(IFS(A10=23,"ok",A12=25,"super"),I3)
