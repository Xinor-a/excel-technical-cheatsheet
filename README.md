# Excel Technical Cheatsheet

## COUNT_CONSECUATIVE_NON_BLANKS

`start`に指定したセルを左上とする行/列全体に対し、`start`から連続して続く値が空でないセルの範囲を求め、その要素数を返します。

<details><summary><b>Formula (行方向):</b></summary>
  
```excel
=LET(
    start,B5,
    r,DROP(INDEX(A:XFD,,COLUMN(start)),ROW(start)-1),
    IFERROR(XMATCH(TRUE,r="")-1,ROWS(r))
)
```

</details>

<details><summary><b>1 Line Formula (行方向):</b></summary>

```excel
=LET(start,B5,r,DROP(INDEX(A:XFD,,COLUMN(start)),ROW(start)-1),IFERROR(XMATCH(TRUE,r="")-1,ROWS(r)))
```

</details>

<details><summary><b>Formula (列方向):</b></summary>
  
```excel
=LET(
    start,B5,
    r,DROP(INDEX(5:5,,COLUMN(start)),,COLUMN(start)-1),
    IFERROR(XMATCH(TRUE,r="")-1,COLUMNS(r))
)
```

</details>

<details><summary><b>1 Line Formula (列方向):</b></summary>

```excel
=LET(start,B5,r,DROP(INDEX(5:5,,COLUMN(start)),,COLUMN(start)-1),IFERROR(XMATCH(TRUE,r="")-1,COLUMNS(r)))
```

</details> 

### 使い方

上記の式を少し改造して、OFFSETを使った連続して値の入ったセルを範囲指定できます。  
`start`に指定したセルを左上とする行/列全体に対し、`start`から連続して続く値が空でないセルの範囲を返します。

<details><summary><b>Formula (行方向):</b></summary>
  
```excel
=LET(
  start,A1,
  cnt,LET(
    r,DROP(INDEX(A:XFD,,COLUMN(start)),ROW(start)-1),
    IFERROR(XMATCH(TRUE,r="")-1,ROWS(r))
  ),
  OFFSET(start,0,0,cnt)
)
```

</details>

<details><summary><b>1 Line Formula (行方向):</b></summary>

```excel
=LET(start,A1,cnt,LET(r,DROP(INDEX(A:XFD,,COLUMN(start)),ROW(start)-1),IFERROR(XMATCH(TRUE,r="")-1,ROWS(r))),OFFSET(start,0,0,cnt))
```

</details>

<details><summary><b>Formula (列方向):</b></summary>
  
```excel
=LET(
  start,A1,
  cnt,LET(
    start,A1,
    r,DROP(INDEX(5:5,,COLUMN(start)),,COLUMN(start)-1),
    IFERROR(XMATCH(TRUE,r="")-1,COLUMNS(r))
  ),
  OFFSET(start,0,0,1,cnt)
)
```

</details>

<details><summary><b>1 Line Formula (列方向):</b></summary>

```excel
=LET(start,A1,cnt,LET(start,A1,r,DROP(INDEX(5:5,,COLUMN(start)),,COLUMN(start)-1),IFERROR(XMATCH(TRUE,r="")-1,COLUMNS(r))),OFFSET(start,0,0,1,cnt))
```

</details>
