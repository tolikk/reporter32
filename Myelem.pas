unit moneytt;

uses Windows;

interface

const OneM: array[1..10] of String = ('', 'iaei ', 'aaa ', 'o?e ', '	aou?a', 'iyou ', 'oanou ', 'naiu ', 'ainaiu ', 'aaayou ');
      OneF: array[1..10] of String = ('', 'iaia ', 'aaa ', 'o?e ', '	aou?a', 'iyou ', 'oanou ', 'naiu ', 'ainaiu ', 'aaayou ');
      Tens: array[1..10] of String = ('', 'aanyou ','aaaaoaou ','o?eaoaou','ni?ie ','iyouaanyo ', 'oanouaanyo ','naiuaanyo ','ainaiuaanyo','aaayiinoi ');
      Ten2: array[1..10] of String = ('','iaeiiaaoaou ','aaaiaaoaou','o?eiaaoaou ','	aou?iaaoaou ','iyoiaaoaou ','oanoiaaoaou ','naiiaaoaou','ainaiiaaoaou ','aaayoiaaoaou ');
      Huns: array[1..10] of String = ('','noi ','aaanoe ','o?enoa','	aou?anoa ','iyounio ','oanounio ','naiunio ','ainaiunio ','aaayounio ');

      Names: array[1..6,1..3] of String = ((' eiiaeea',' eiiaeee','eiiaae'),
                                    ('?oaeu ','?oaey ','?oaeae '),
                                    ('ouny	a ','ouny	e ','ouny	 '),
                                    ('ieeeeii ','ieeeeiia ','ieeeeiiia '),
                                    ('ieeeea?a ','ieeeea?aa ','ieeeea?aia'),
                                           ('o?eeeeii ','o?eeeeiia','o?eeeeiiia '));

      NamesISO: array[1..6,1..3] of String = (('','',''),
                                    ('','',''),
                                    ('ouny	a ','ouny	e ','ouny	 '),
                                    ('ieeeeii ','ieeeeiia ','ieeeeiiia '),
                                    ('ieeeea?a ','ieeeea?aa ','ieeeea?aia'),
                                           ('o?eeeeii ','o?eeeeiia','o?eeeeiiia '));



                                           //type
function UpperRus( var b : char ): char;
function DigToStr(var buff : char; dig : double; level : integer; var ISO :char):char;

implementation

function UpperRus( var b : char  ): char;
begin
 case( b ) of
  'e': b := 'E';
  'o': b := 'O';
  'o': b := 'O';
  'e': b := 'E';
  'a': b := 'A';
  'i': b := 'I';
  'a': b := 'A';
  'o': b := 'O';
  'u': b := 'U';
  'c': b := 'C';
  'o': b := 'O';
  'u': b := 'U';
  'o': b := 'O';
  'u': b := 'U';
                'a': b := 'A';
  'a': b := 'A';
  'i': b := 'I';
  '?': b := '?';
  'i': b := 'I';
  'e': b := 'E';
  'a': b := 'A';
  '?': b := '?';
  'y': b := 'Y';
  'y': b := '?';
  '	': b := '?';
  'n': b := 'N';
  'i': b := 'I';
  'e': b := 'E';
  'o': b := 'O';
  'u': b := 'U';
  'a': b := 'A';
  '?': b := '?';
  '?': b := '?';
       end;
end;

function DigToStr(var buff : string; dig : double ; level : integer ; var ISO : char):string;
var
  tmpstr  : string[100] ;
  tmpstr2 : string[100] ;
begin
  tmpstr2[1]:= '0';
  tmpstr[1]:= '0';

  if (level >0) then
    if (((dig - (dig / 100) * 100) >= 11) and
        ((dig - (dig / 100) * 100) <= 19)) then
        tmpstr := Ten2[trunc((dig - (dig / 100) * 100) - 10)]
    else begin
      if (level = 2) then
  tmpstr := OneF[trunc(dig - (dig / 10) * 10)]
  else
  tmpstr := OneM[trunc(dig - (dig / 10) * 10)];
  tmpstr2:= tmpstr;
  tmpstr := Tens[trunc((dig / 10) - (dig / 100) * 10)];
  tmpstr := tmpstr + tmpstr2;
  end;
  tmpstr2 :=  tmpstr;
  tmpstr := Huns[trunc(dig / 100)];
  tmpstr := tmpstr + tmpstr2;
 // end;

  if (((dig - (dig / 100) * 100) >= 11) and ((dig - (dig / 100) * 100) <=
19)) then
       tmpstr := tmpstr + Names[level][2]
  else
  case (trunc(dig - (dig / 10) * 10)) of
   0: tmpstr := tmpstr + Names[level][3];
                 5: tmpstr := tmpstr + Names[level][3];
                 6: tmpstr := tmpstr + Names[level][3];
                 7: tmpstr := tmpstr + Names[level][3];
                 8: tmpstr := tmpstr + Names[level][3];
                 9: tmpstr := tmpstr + Names[level][3];
   1: tmpstr := tmpstr + Names[level][1];
                 2: tmpstr := tmpstr + Names[level][2];
                 3: tmpstr := tmpstr + Names[level][2];
                 4: tmpstr := tmpstr + Names[level][2];
//        strcat(tmpstr, (ISO?NamesISO[level][1]:Names[level][1])); break;
    end;

  DigToStr:= tmpstr;
  //return buff;
end;



end.
