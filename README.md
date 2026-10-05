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
