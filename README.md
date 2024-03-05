

```
{column_name} * 2
```

```
{column_name[3]} * 2
```

```
{column_name[:]} * 2
```

```
{column_name$[2]} * 2
```

```
{column_name$[2]} * 2
```

```
{tab_name:column_name} * 2
```

```
{sheet_name:tab_name:column_name} * 2
```

```
{sheet_name:tab_name:column_name} * 2
```

```
{column_name:SHIFT(-1,1)} * 2
```

# Idea: for each table, for each formula, run through recursively until you determine the maximum size of the table.

```
{sheet_name:tab_name:column_name:SHIFT(-1,1)}
```