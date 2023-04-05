// 選択範囲のSQL文を整形
var
    mess:String;            // メッセージ出力用変数
    org_str,target,w_str:String;    // 文字保存用変数
    tmp1,tmp2,w_tmp:String;     // 文字判定用ワーク変数
    CRLF,TAB:String;        // 制御コード保持用変数
    byte:Integer;           // 変換対象文字列サイズ
    pt:Integer;         // 変換対象文字列ポインター
    t_flg:Integer;          // インデント数
    i_cnt:Integer;          // インデント挿入用カウンター
    f_flg:Integer;          // '('のあと副問い合わせがある場合')'でインデント数を-2するフラグ
    s_flg:Integer;          // '('の前にキーワードがある場合改行するフラグ
 
begin
    CRLF  := '
';
    TAB  := '    ';
    org_str := S_GetSelectedString(0);
    byte := Length(org_str);
    if (byte=0) then
        begin
            S_SelectAll();
            org_str := S_GetSelectedString(0);
            byte := Length(org_str);
        end
    ;
    pt := 1;
    t_flg := 0;
    f_flg := 0;
    s_flg := 0;
 
    target := '';
    w_str := '';
 
    while (pt <= byte) do
    begin
        tmp1 := Copy(org_str,pt,1);
        tmp2 := Copy(org_str,pt,2);
        if ((tmp1=' ') or (tmp1=TAB) or (tmp1=',') or (tmp2 = CRLF) or (tmp1='(') or (tmp1=')') or (tmp1 = ';') ) then
            begin
                w_tmp := LowerCase(Trim(w_str));
                if ((w_tmp='select') or (w_tmp='update') or (w_tmp='commit') or (w_tmp='rollback') or (w_tmp='insert') or (w_tmp='delete') or (w_tmp='create') or (w_tmp='drop') or (w_tmp='truncate') or (w_tmp='alter')) then
                    begin
                        s_flg := 0;
                        if (f_flg<>0) then f_flg := 2;
                        i_cnt := 0;
                        while (i_cnt < t_flg) do
                        begin
                            target := target + TAB;
                            i_cnt := i_cnt + 1;
                        end;
                        if ((w_tmp='select') or (w_tmp='update') or (w_tmp='commit') or (w_tmp='rollback')) then
                            begin
                                t_flg := t_flg + 1;
                                target := target + Trim(w_str) + CRLF;
                            end
                        else
                            begin
                                target := target + Trim(w_str) + ' ';
                            end
                        ;
                    end
                else if ((w_tmp='values') or (w_tmp='set') or (w_tmp='from') or (w_tmp='where') or (w_tmp='having') or (w_tmp='group') or (w_tmp='order')) then
                    begin
                        if (w_tmp='values') then
                            begin
                                 s_flg := 2;
                             end
                         else
                             begin
                                 s_flg := 0;
                                target := target + CRLF;
                             end
                         ;
                        if (f_flg<>0) then f_flg := 2;
                        t_flg := t_flg - 1;
                        i_cnt := 0;
                        while (i_cnt < t_flg) do
                        begin
                            target := target + TAB;
                            i_cnt := i_cnt + 1;
                        end;
                        if ((w_tmp='group') or (w_tmp='order')) then
                            begin
                                target := target + Trim(w_str) + ' ';
                            end
                        else
                            begin
                                t_flg := t_flg + 1;
                                target := target + Trim(w_str) + CRLF;
                            end
                        ;
                    end
                else if ((w_tmp='and') or (w_tmp='or') or (w_tmp='table') or (w_tmp='view') or (w_tmp='index') or (w_tmp='by') or (w_tmp='into')) then
                    begin
                        s_flg := 0;
                        if ((w_tmp='table') or (w_tmp='view') or (w_tmp='index')) then
                            begin
                                target :=target + Trim(w_str) + ' ';
                            end
                        else if ((w_tmp='by') or (w_tmp='into')) then
                            begin
                                target := target + Trim(w_str) + CRLF;
                                t_flg := t_flg + 1;
                            end
                        else
                            begin
                                target :=target + ' ' + Trim(w_str) + CRLF;
                            end
                        ;
                    end
                else if (tmp1=',') then
                    begin
                        s_flg := 2;
                        i_cnt := 0;
                        while (i_cnt < t_flg) do
                        begin
                            target := target + TAB;
                            i_cnt := i_cnt + 1;
                        end;
                        target := target + Trim(w_str) + tmp1 + CRLF;
                    end
                else if (tmp1='(') then
                    begin
                        f_flg := 1;
                        if (s_flg=1) then target := target + CRLF;
                        if (Length(w_tmp) <> 0) then
                            begin
                                s_flg := 1;
                                i_cnt := 0;
                                while (i_cnt < t_flg) do
                                begin
                                    target := target + TAB;
                                    i_cnt := i_cnt + 1;
                                end;
                                target := target + Trim(w_str) + CRLF;
                            end
                        ;
                        i_cnt := 0;
                        while (i_cnt < t_flg) do
                        begin
                            target := target + TAB;
                            i_cnt := i_cnt + 1;
                        end;
                        t_flg := t_flg + 1;
                        target := target + tmp1 + CRLF;
                    end
                else if (tmp1=')') then
                    begin
                        if (Length(w_tmp) <> 0) then
                            begin
                                i_cnt := 0;
                                while (i_cnt < t_flg) do
                                begin
                                    target := target + TAB;
                                    i_cnt := i_cnt + 1;
                                end;
                            end
                        ;
                        target := target + Trim(w_str) + CRLF;
// mess := 'f_flg[' +inttostr(f_flg) + ']' + target + CRLF;
// MessageBox(mess,'選択中の文字列',0);
                        if (f_flg=2) then t_flg := t_flg - 2 else t_flg := t_flg - 1;
                        i_cnt := 0;
                        while (i_cnt < t_flg) do
                        begin
                            target := target + TAB;
                            i_cnt := i_cnt + 1;
                        end;
                        target := target + tmp1 + CRLF;
                        if (t_flg<0) then t_flg := 0;
                        f_flg := 0;
                    end
                else if (tmp1=';') then
                    begin
                        if (Length(w_tmp) <> 0) then
                            begin
                                i_cnt := 0;
                                while (i_cnt < t_flg) do
                                begin
                                    target := target + TAB;
                                    i_cnt := i_cnt + 1;
                                end;
                            end
                        ;
                        t_flg := 0;
                        f_flg := 0;
                        s_flg := 0;
                        target := target + Trim(w_str) + ';' + CRLF;
                    end
                else
                    begin
                        if (Length(w_tmp) <> 0) then
                            begin
                                i_cnt := 0;
                                while (i_cnt < t_flg) do
                                begin
                                    target := target + TAB;
                                    i_cnt := i_cnt + 1;
                                end;
                            end
                        ;
                        if (tmp2 = CRLF) then
                            begin
                                target := target + Trim(w_str);
                                if (Length(w_tmp) <> 0) then s_flg := 1;
                            end
                        else
                            begin
                                if ((tmp1<>TAB) and (tmp1<>' ')) then
                                    begin
                                        target := target + Trim(w_str) + tmp1;
                                        s_flg := 0;
                                    end
                                else
                                    begin
                                        target := target + Trim(w_str);
                                        if (s_flg<>2) then s_flg := 1;
                                    end
                                ;
                            end
                        ;
                    end
                ;
                w_str := '';
            end
        else
            begin
                w_str := w_str + tmp1;
            end
        ;
        pt := pt + 1;
    end;
    if (w_str<>'') then target := target + Trim(w_str);
    target := target + CRLF;
    S_Delete();
    S_InsText(target);
end;