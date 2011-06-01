unit MandReservDlgUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, RXSpin, ComCtrls, Db, DBTables, RxLookup, Placemnt, Mask;

type
  TMandatorReserv = class(TForm)
    Button1: TButton;
    Label1: TLabel;
    ReportDate: TDateTimePicker;
    Label2: TLabel;
    FondPercent: TRxSpinEdit;
    Label3: TLabel;
    FizichTax: TRxSpinEdit;
    Label4: TLabel;
    UridichTax: TRxSpinEdit;
    IsAgent: TCheckBox;
    Agents: TQuery;
    DataSourceAgents: TDataSource;
    AgentCombo: TRxDBLookupCombo;
    CalcBtn: TButton;
    WorkSQL: TQuery;
    StatusText: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure CalcBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MandatorReserv: TMandatorReserv;

implementation

uses MF;

{$R *.DFM}

procedure TMandatorReserv.FormCreate(Sender: TObject);
begin
     Agents.Open
end;

//Коды состояний полисов
//0 "";
//1 "ИСПОРЧЕН";
//2 "ДОСРОЧНО ЗАВЕРШЁН";
//3 "УТЕРЯН";
//4 "СРОК ДЕЙСТВИЯ ЗАКОНЧИЛСЯ";
//5 "ЗАВЕРШЁН ПО НЕУПЛАТЕ";

procedure TMandatorReserv.CalcBtnClick(Sender: TObject);
var
    s : string;
    i, j : integer;
    Summs : array [1..10] of double;
    Currs : array [1..10] of string;
begin
     {
     for i := 1 to 10 do begin
         Summs[i] := 0;
         Currs[i] := 'TOLIK';
     end;

     if IsAgent.Checked then
          if AgentCombo.Text = '' then begin
             AgentCombo.DropDown;
             exit;
          end;

    WorkSQL.Close;
    WorkSQL.SQL.Clear;
    WorkSQL.SQL.Add('SELECT * FROM MANDATOR WHERE PAY1DT IS NOT NULL AND (PSER IS NULL OR PSER='') AND (PAY1DT <= :REPDT AND TODT >= :REPDT))');
    WorkSQL.ParamByName('REPDT').AsDate := ReportDate.Date;
    WorkSQL.Open;
			while(not WorkSQL.Eof) do begin
				sqldate stopdate = getPolises["STOPDATE"];
				if(!stopdate.is_null())
					if(getDate(stopdate).m_dt <= dlg.m_RepDate.m_dt)
						continue;

				if(f)
					fprintf(f, "%s/%lu ", (LPCTSTR)getPolises["SER"], (int)getPolises["NMB"]);

				PolisCounter++;

				if((PolisCounter % 100) == 0) begin
					m_wndStatusBar.SetWindowText("Обработано " + LongToString(PolisCounter));
					m_wndStatusBar.UpdateWindow();
				end

				CBlankString PAYCURR(getPolises["PAY1CURR"]);
				ReportStructure* dataptr = 0;
				for(int f_i = 0; f_i < 10; f_i++) begin
					if( *rs[f_i].Curr == 0 || PAYCURR == rs[f_i].Curr) begin
						dataptr = &rs[f_i];
						strcpy(dataptr->Curr, PAYCURR);
						break;
					end
				end
				ASSERT(dataptr);

				sqldate todt = getPolises["TODT"];
				sqldate frdt = getPolises["FRDT"];
				//sqldate pay1dt = getPolises["PAY1DT"];
				int AllPeriod = getDate(todt).m_dt - getDate(frdt).m_dt + 1; 

				if(f)
					fprintf(f, "до %s ", (const char*)todt.to_string());

				if(PAYCURR.IsEmpty()) 
					throw xsql("Не определена валюта оплаты " + CBlankString(getPolises["SER"]) + '/' + LongToString(getPolises["NMB"]));
					
				if(PAYCURR != "BRB") begin
					//Валюта
					//валютное отчисление в фонд
					double BruttoPremium = getPolises["PAY1"];
					double Otchislen = BruttoPremium * dlg.GetFondPercent(getPolises["PAY1DT"], (LPCTSTR)getPolises["SER"], getPolises["NMB"]) / 100;
					dataptr->FOND_PREMUIM += Otchislen;

					if(f) 
						fprintf(f, "1пл. %0.2f %s, ", BruttoPremium, PAYCURR);

					//% продажи
					double PercentPay = dlg.GetSellPercent(getPolises["PAY1DT"], (LPCTSTR)getPolises["SER"], getPolises["NMB"]) / 100;

					if(f) 
						fprintf(f, "Продажа %0.2f%c, ", PercentPay * 100, '%');

					double BRBSum = BruttoPremium * PercentPay;
					double AgentSumma = BruttoPremium * (double)getPolises["AGPCNT"] / 100;
					CBlankString AgType = getPolises["AGTYPE"];
					if(AgType.GetLength() != 1 || CBlankString("FU").Find(AgType[0]) == -1)
						throw xsql("Неизвестный тип агента " + CBlankString(getPolises["SER"]) + '/' + LongToString(getPolises["NMB"]));

					if(AgType[0] == 'U') 
						AgentSumma = AgentSumma * (1 + dlg.m_UrTax / 100);
					else
						AgentSumma = AgentSumma * (1 + dlg.m_FizTax / 100);

					if(f) 
						fprintf(f, "Агент %0.2f, ", AgentSumma);

					BruttoPremium *= (1 - PercentPay);
					dataptr->BRUTTO_PREMUIM += BruttoPremium;

					double bp = BruttoPremium - Otchislen;
					dataptr->BASE_PREMUIM += bp;

					if(f) 
						fprintf(f, "БП %0.2f, ", bp);

					if(1 /*AllPeriod > 30*/) begin
						int WorkPeriod = getDate(todt).m_dt - dlg.m_RepDate.m_dt; 
						double rnp = WorkPeriod < 0 ? bp : (bp * max(0, WorkPeriod) / AllPeriod);
						dataptr->RNP2_PREMUIM += rnp;

						if(f) 
							fprintf(f, "РНП %0.2f, ", rnp);
					end

					SQLQuery MaxDate(RateCur, "SELECT MAX(RATEDATE) FROM SYSADM.RATES WHERE RATEDATE<=:1");
					MaxDate << sqldate(getPolises["PAY1DT"]);
					if(!MaxDate.StartFetch())
						throw xsql("Нет курса валюты на дату оплаты " + CBlankString(getPolises["SER"]) + '/' + LongToString(getPolises["NMB"]));

					sqldate date = MaxDate[1];
		
					SQLQuery getRate(RateCur, "SELECT " + PAYCURR + "_NAL FROM SYSADM.RATES WHERE RATEDATE=:1");
					getRate << date;
					if(!getRate.StartFetch())
						throw xsql("Нет курса валюты на дату оплаты " + CBlankString(getPolises["SER"]) + '/' + LongToString(getPolises["NMB"]));
					
					double Rate = getRate[1];
					if(f) 
						fprintf(f, "КУРС %0.2f, ", Rate);

					CBlankString PAYCURR = "BRB";
					dataptr = 0;
					for(int f_i = 0; f_i < 10; f_i++) begin
						if( *rs[f_i].Curr == 0 || PAYCURR == rs[f_i].Curr) begin
							dataptr = &rs[f_i];
							strcpy(dataptr->Curr, PAYCURR);
							break;
						end
					end
					ASSERT(dataptr);

					if(CBlankString(getPolises["ISPAY2"]) == "Y" && CBlankString(getPolises["PAY2CURR"]) != "BRB") 
						throw xsql("Валютная оплата в 2 этапа " + CBlankString(getPolises["SER"]) + '/' + LongToString(getPolises["NMB"]));

					bp = 0.;
					if(CBlankString(getPolises["ISPAY2"]) == "Y") begin
						double BruttoPremium = (double)getPolises["PAY2"];
						double Otchislen = BruttoPremium * dlg.GetFondPercent(getPolises["PAY2DT"], (LPCTSTR)getPolises["SER"], getPolises["NMB"]) / 100.;
						double AgSumma = BruttoPremium * (double)getPolises["AGPCNT"] / 100.;
						if(AgType[0] == 'U') 
							AgSumma = AgSumma * (1 + dlg.m_UrTax / 100);
						else
							AgSumma = AgSumma * (1 + dlg.m_FizTax / 100);
						dataptr->BRUTTO_PREMUIM += BruttoPremium; //полученные рубли
						dataptr->FOND_PREMUIM += Otchislen; //отчисление в фонд
						dataptr->AGENT_PREMUIM += AgSumma; //процент агенту
						bp = BruttoPremium - Otchislen - AgSumma;
						dataptr->BASE_PREMUIM += bp; //базовая

						if(f) 
							fprintf(f, "2й пл. %0.2f, ", BruttoPremium);
						if(f) 
							fprintf(f, "Фонды %0.2f, ", Otchislen);
						if(f) 
							fprintf(f, "Агент %0.2f, ", AgSumma);
						if(f) 
							fprintf(f, "БП %0.2f, ", bp);
					end

					dataptr->BRUTTO_PREMUIM += BRBSum * Rate;
					dataptr->AGENT_PREMUIM += AgentSumma * Rate;
					bp += (BRBSum - AgentSumma) * Rate;
					dataptr->BASE_PREMUIM += bp;

					if(1/*AllPeriod > 30*/) begin
						int WorkPeriod = getDate(todt).m_dt - dlg.m_RepDate.m_dt; 
						double rnp = (dlg.m_RepDate.m_dt < getDate(frdt).m_dt) ? bp : (bp * max(0, WorkPeriod) / AllPeriod);
						dataptr->RNP2_PREMUIM += rnp;
						if(f) 
							fprintf(f, "РНП %0.2f, ", rnp);
					end
				end

				if(PAYCURR == "BRB") begin
					double BruttoPremium = getPolises["PAY1"];
					double BruttoPremium1 = getPolises["PAY1"];
					double BruttoPremium2 = 0;
					double FondPercent2 = 0;

					if(f)
						fprintf(f, "1пл. %0.2f, ", BruttoPremium);

					double PAY2DUP = 0;

					if(int(getPolises["STATE"]) == STATE_POLIS_LOST) //Утерян
						if(CBlankString(getPolises["ISFEE2"]) == "Y") //Нужна оплата
							if(CBlankString(getPolises["ISPAY2"]) == "N") begin //Нет оплаты
								sqldate fee2date = getPolises["FEE2DT"];
								if(fee2date.is_null())
									throw xsql("Нет даты 2й оплаты " + CBlankString(getPolises["SER"]) + '/' + LongToString(getPolises["NMB"]));
								if(getDate(fee2date).m_dt < dlg.m_RepDate.m_dt) begin //Должна быть!!!
									SQLQuery getDupPolis(RateCur, "SELECT * FROM SYSADM.MANDATOR WHERE PSER=:1 AND PNMB=:2");
									getDupPolis << CBlankString(getPolises["SER"])
										        << int(getPolises["NMB"]);
									if(getDupPolis.StartFetch()) begin //Есть дубликат
										if(int(getDupPolis["STATE"]) != STATE_POLIS_STOPPAY) begin
											if(CBlankString(getDupPolis["ISPAY2"]) == "Y") begin
												PAY2DUP = getDupPolis["PAY2"];
												FondPercent2 = dlg.GetFondPercent(getDupPolis["PAY2DT"], (LPCTSTR)getDupPolis["SER"], getDupPolis["NMB"]) / 100;
												if(PAY2DUP < 0.01)
													throw xsql("Нет 2й оплаты " + CBlankString(getDupPolis["SER"]) + '/' + LongToString(getDupPolis["NMB"]));
											end
											else
											if(int(getDupPolis["STATE"]) == STATE_POLIS_LOST)
												NotSupportPolis++;
										end
									end
									else
										NotFoundDup++;
								end
							end

					//check 2 pay
					if(CBlankString(getPolises["ISPAY2"]) == "Y") begin
						if(dlg.m_RepDate.m_dt > getDate(getPolises["PAY2DT"])) begin
							if(CBlankString(getPolises["PAY2CURR"]) != "BRB")
								throw xsql("Проблема со 2й оплатой (она валютная) " + CBlankString(getPolises["SER"]) + '/' + LongToString(getPolises["NMB"]));
							BruttoPremium += (double)getPolises["PAY2"];
							BruttoPremium2 = (double)getPolises["PAY2"];
							FondPercent2 = dlg.GetFondPercent(getPolises["PAY2DT"], (LPCTSTR)getPolises["SER"], getPolises["NMB"]) / 100;
						end
					end

					if(BruttoPremium2 < 0.01) begin
						BruttoPremium2 = PAY2DUP;
						BruttoPremium += PAY2DUP;
					end

					if(f && BruttoPremium2 > 1)
						fprintf(f, "2пл. %0.2f, ", BruttoPremium2);

					dataptr->BRUTTO_PREMUIM += BruttoPremium;

					double AgentSumma = BruttoPremium * (double)getPolises["AGPCNT"] / 100;
					CBlankString AgType = getPolises["AGTYPE"];
					if(AgType.GetLength() != 1 || CBlankString("FU").Find(AgType[0]) == -1)
						throw xsql("Неизвестный тип агента " + CBlankString(getPolises["SER"]) + '/' + LongToString(getPolises["NMB"]));

					if(AgType[0] == 'U')
						AgentSumma = AgentSumma * (1 + dlg.m_UrTax / 100);
					else
						AgentSumma = AgentSumma * (1 + dlg.m_FizTax / 100);

					dataptr->AGENT_PREMUIM += AgentSumma;

					if(f && AgentSumma > 1)
						fprintf(f, "агент %0.2f, ", AgentSumma);

					double Otchislen = BruttoPremium1 * dlg.GetFondPercent(getPolises["PAY1DT"], (LPCTSTR)getPolises["SER"], getPolises["NMB"]) / 100;
					if(BruttoPremium2 > 0)
						Otchislen += BruttoPremium2 * FondPercent2;// dlg.GetFondPercent(getPolises["PAY2DT"], (LPCTSTR)getPolises["SER"], getPolises["NMB"]) / 100;
					dataptr->FOND_PREMUIM += Otchislen;

					if(f && Otchislen > 1)
						fprintf(f, "фонды %0.2f, ", Otchislen);

					double bp = (BruttoPremium - AgentSumma - Otchislen);
					dataptr->BASE_PREMUIM += bp;

					if(f)
						fprintf(f, "БП %0.2f, ", bp);

					if(1/*AllPeriod > 30*/) begin
						int WorkPeriod = getDate(todt).m_dt - dlg.m_RepDate.m_dt + 1;
						double rnp = (dlg.m_RepDate.m_dt < getDate(frdt).m_dt) ? bp : (bp * max(0, WorkPeriod) / AllPeriod);
						dataptr->RNP2_PREMUIM += rnp;
						if(f)
							fprintf(f, "РНП2 %lu/%lu = %0.2f\r", WorkPeriod, AllPeriod, rnp);
					end
				end
			end
     }       
     Screen.Cursor := crDefault;
end;

procedure TMandatorReserv.FormShow(Sender: TObject);
begin
          StatusText.Text := '';
end;

end.
