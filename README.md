# Excel Technical Cheatsheet

## COUNT_CONSECUATIVE_NON_BLANKS

`start`に指定したセルを左上とする行/列全体に対し、`start`から連続して続く値が空でないセルの範囲を求め、その要素数を返します。

<details><summary><b>COUNT_CONSECUATIVE_NON_BLANKS_ROW</b></summary>
  
  ```excel
  =LAMBDA(start,LET(r,DROP(INDEX(A:XFD,,COLUMN(start)),ROW(start)-1),IFERROR(XMATCH(TRUE,r="")-1,ROWS(r))))
  ```

</details>

<details><summary><b>COUNT_CONSECUATIVE_NON_BLANKS_COLUMN</b></summary>
  
  ```excel
  =LAMBDA(start,LET(r,DROP(INDEX(5:5,,COLUMN(start)),,COLUMN(start)-1),IFERROR(XMATCH(TRUE,r="")-1,COLUMNS(r))))
  ```

</details> 

## RANGE_CONSECUATIVE_NON_BLANKS

[COUNT_CONSECUATIVE_NON_BLANKS](#COUNT_CONSECUATIVE_NON_BLANKS)を使って、OFFSETで連続して値の入ったセルを範囲指定できます。  
`start`に指定したセルを左上とする行/列全体に対し、`start`から連続して続く値が空でないセルの範囲を返します。

<details><summary><b>RANGE_CONSECUATIVE_NON_BLANKS_ROW</b></summary>

  ```excel
  =LAMBDA(start,OFFSET(start,0,0,COUNT_CONSECUATIVE_NON_BLANKS_ROW(start)))
  ```

</details>

<details><summary><b>RANGE_CONSECUATIVE_NON_BLANKS_COLUMN</b></summary>

  ```excel
  =LAMBDA(start,OFFSET(start,0,0,1,COUNT_CONSECUATIVE_NON_BLANKS_COLUMN(start)))
  ```

</details>
