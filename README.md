A VBScript that exercises the DuckDB ODBC driver using ADO. 

DuckDB ODBC driver: https://github.com/duckdb/duckdb-odbc/releases

To run, open a cmd prompt in this folder and run

```
runOdbcTest.bat
```

Since DuckDB is currently 64 bit only, the .bat/.wsf are set up to ensure the script is run 64-bit.

An example of the output of the script is included in sample_output.txt. Your output may vary slightly depending on which version of the driver you have installed. In the case here, the version on github is identified as 1.2.1, however the `PRAGMA version;` reports v1.1.4-dev1904,fb7701fec0.

