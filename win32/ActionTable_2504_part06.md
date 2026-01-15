# ActionTable_2504_part06

## Page 51

### 예제 : 배포용 문서 만들기 (계속)

```javascript
        var msg = "배포용 문서 만들기 실패";
        if (pHwpCtrl.EditMode & 0x10) // 배포용 문서는 0x10 flag 를 포함한다.
            msg += "\n이미 배포용 문서로 지정된 상태입니다.\n암호를 변경하기 위해서는 먼저 일반 문서로 변경하십시오."
        else if (pHwpCtrl.EditMode == 0)
            msg += "\n읽기 전용 문서입니다."
        alert(msg);
    }
    pHwpCtrl.focus();
}
```

### 예제 : 언더라인이 있는 문자열의 시작 pos를 얻어오는 예제

```javascript
function OnTestApi1()
{
    pHwpCtrl.MovePos(3);                                    // 커서를 문서의 맨 뒤로 이동
    var act = pHwpCtrl.CreateAction("BackwardFind");        // 뒤에서부터 찾는다.
    var set = act.CreateSet();
    act.GetDefault(set);

    set.SetItem("IgnoreFindString", 1);
    var subset = set.CreateItemSet("FindCharShape", "CharShape");
    subset.SetItem("UnderlineType", 1);
    act.Execute(set);
    var set;
    var list, para, pos;
    set = pHwpCtrl.GetPosBySet();
    list = set.Item("List");
    para = set.Item("Para");
    pos = set.Item("Pos");
}
```

### 예제 : 각 라인의 첫 번째 단어의 색깔을 바꾸는 예제

```javascript
function OnTestApi1()
{
    // 현재줄의 맨 앞으로 커서를 옮긴후 한단어를 블럭지정한다.
    pHwpCtrl.Run("MoveLineBegin");
    pHwpCtrl.Run("Select");
    pHwpCtrl.Run("MoveSelNextWord");
```

## Page 52

```javascript
    // 블럭 지정된 단어의 글자 속성을 변경한다.
    var dAct = pHwpCtrl.CreateAction("CharShape");
    var dSet = dAct.CreateSet();
    dAct.GetDefault(dSet);
    dSet.SetItem("TextColor", 0xFF0000);     // 글자 색을 파란색으로
    dAct.Execute(dSet);

    pHwpCtrl.Run("Cancel");
}

function OnTestApi2()
{
    // 문서의 맨 처음에서 끝까지 돌면서 각 줄의 맨 처음 단어의 글자속성을 변경한다.
    pHwpCtrl.MovePos(2);                     // 페이지 맨 처음으로
    var con = true;

    while (con) {
        OnTestApi1();                        // 현재줄의 맨 처음 단어 색깔 변경
        con = pHwpCtrl.MovePos(20);          // 한줄 아래로 (예전 API 매뉴얼에는 21로 되어 있으나 잘못된 값임.
                                             // 한줄 위/아래로가 서로 값이 바뀌었음)
    }
}
```
