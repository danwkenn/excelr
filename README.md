

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
{sheet_name:tab_name:column_name:SHIFT}
```

```
{S-sheet_name:T-tab_name:F-field_name:C-column_name}
```

```
{a} -> {S-this_sheet:T-this_tab:F-body:C-a}
```

```
{F-b:a} -> {S-this_sheet:T-this_tab:F-b:C-a}
```

```
{F-b:C-a} -> {S-this_sheet:T-this_tab:F-b:C-a}
```

```
{c:F-b:a} -> {S-this_sheet:T-c:F-b:C-a}
```

```
{c:b:a}/{S-c:T-b:C-a} -> {S-c:T-b:F-body:C-a}
```

```
{d:c:b:a} -> {S-d:T-c:F-b:C-a}
```

```
{d:c:b:a} -> {S-d:T-c:F-b:C-a}
```