"""Fixed-value Excel 2010 verification inputs, including in-place Excel edits."""
from dataclasses import dataclass, field, replace
from datetime import date
from pathlib import Path
import argparse
import hashlib
import json
from openpyxl import load_workbook
from src.lifeplan import demo, Rule, Account
from src.p2a import CashEvent, from_p1, calculate
from src.p2a_excel import generate, read_plan, identity


def one(**changes):
    args=dict(id='E001',name='旅行',kind='expense',amount=1000000,account_id='cash',occurrence_month=(2026,7),additional_confirmed=True)
    args.update(changes)
    return CashEvent(**args)


def cash_plan(start=(2026,1),end=(2026,12),cash=1000000):
    p=from_p1(demo())
    return replace(p,start=start,end=end,rules=[],accounts=[replace(p.accounts[0],balance=cash)])


@dataclass
class Case:
    name: str
    plan: object
    expected_cash: int | None = None
    edits: dict = field(default_factory=dict)
    invalid: bool = False
    monitor_seconds: int = 30


def cases():
    p=from_p1(demo())
    yield Case('baseline',p,2200000)
    salary,living=p.rules
    changed=replace(p,rules=[replace(salary,end=(2026,6)),replace(salary,id='raise',annual=4800000,start=(2026,7)),living])
    yield Case('income_change',changed,2800000)
    yield Case('partial_income',replace(changed,start=(2026,4),end=(2026,9)),1900000)
    yield Case('once_expense',replace(p,events=[one()]),1200000)
    yield Case('once_income',replace(p,events=[one(kind='income',amount=600000,occurrence_month=(2026,11))]),2800000)
    yield Case('once_100',replace(cash_plan(),events=[one(amount=100,occurrence_month=(2026,3))]),999900)
    cycle=one(event_type='periodic',occurrence_month=None,first_month=(2030,5),last_month=(2042,5),interval_months=48,amount=3000000)
    q=replace(cash_plan((2029,1),(2043,12),20000000),events=[cycle])
    yield Case('cycle_48',q,8000000)
    yield Case('cycle_end',replace(q,events=[replace(cycle,last_month=(2042,4))]),11000000)
    yield Case('cycle_anchor',replace(q,start=(2032,1)),11000000)
    yield Case('cycle_14',replace(cash_plan((2026,1),(2030,12)),events=[one(event_type='periodic',occurrence_month=None,first_month=(2026,11),last_month=(2029,12),interval_months=14,amount=100)]),999700)
    q=replace(p,events=[one(occurrence_month=(2026,3)),one(id='E002',kind='income',amount=600000,occurrence_month=(2026,10))],start=(2026,4),end=(2026,9))
    yield Case('outside',q,1600000)
    q=replace(p,rules=[living],accounts=[p.accounts[0],replace(p.accounts[1],balance=1000000)],events=[one(occurrence_month=(2026,3))])
    yield Case('protected_shortage',q,-2400000)
    yield Case('zero_balance',replace(cash_plan(),events=[one(occurrence_month=(2026,3))]),0)
    q=replace(cash_plan((2026,1),(2028,12)),events=[one(occurrence_month=(2028,1),inflation_enabled=True,base_month=(2026,1),rate_bps=200)])
    yield Case('inflation_2',q,-40400)
    yield Case('inflation_small',replace(q,events=[replace(q.events[0],amount=6,rate_bps=1000)]),999992)
    yield Case('deflation',replace(q,events=[replace(q.events[0],rate_bps=-1000)]),190000)
    yield Case('inflation_cycle',replace(cash_plan((2030,1),(2034,12)),events=[one(event_type='periodic',occurrence_month=None,first_month=(2030,5),last_month=(2034,5),interval_months=48,inflation_enabled=True,base_month=(2030,5),rate_bps=200)]),-1082432)
    template=replace(p,events=[one()])
    added={'イベント入力':{'B3':'家電','E3':date(2026,8,1),'H3':100000,'I3':'cash','M3':'はい'}}
    yield Case('excel_add',template,1100000,added)
    yield Case('excel_change',template,1700000,{'イベント入力':{'H2':500000}})
    yield Case('excel_disable',template,2200000,{'イベント入力':{'S2':0}})
    yield Case('excel_clear',template,2200000,{'イベント入力':{f'{chr(c)}2':None for c in range(ord('B'),ord('S')+1)}})
    yield Case('excel_income_change',p,2800000,{'収支':{'F2':date(2026,6,1),'B4':'salary','C4':'income','D4':4800000,'E4':date(2026,7,1),'F4':date(2026,12,1),'G4':'cash','T4':'salary'}})
    inflated=replace(template,events=[replace(template.events[0],inflation_enabled=True,base_month=(2024,1),rate_bps=200)])
    yield Case('excel_rate_change',inflated,1139100,{'イベント入力':{'P2':300}})
    bad={'B3':'同じ費用','E3':date(2026,7,1),'H3':100000,'I3':'cash','M3':'はい','L3':'E001'}
    yield Case('duplicate_event',template,edits={'イベント入力':bad},invalid=True)
    yield Case('duplicate_id',template,edits={'イベント入力':{**bad,'A3':'E001','L3':'OTHER'}},invalid=True)
    yield Case('duplicate_rule',template,edits={'イベント入力':{'L2':'salary'}},invalid=True)
    yield Case('missing_confirmation',template,edits={'イベント入力':{'M2':None}},invalid=True)
    yield Case('zero_interval',template,edits={'イベント入力':{'C2':'周期','F2':date(2030,7,1),'G2':0}},invalid=True)
    yield Case('invalid_rate',template,edits={'イベント入力':{'N2':1,'O2':date(2026,1,1),'P2':1001}},invalid=True)
    yield Case('price_overflow',template,edits={'イベント入力':{'N2':1,'O2':date(2025,1,1),'P2':1,'H2':1000000000}},invalid=True)
    yield Case('income_inflation',template,edits={'イベント入力':{'D2':'収入','N2':1,'O2':date(2026,1,1),'P2':200}},invalid=True)
    yield Case('partial_row',template,edits={'イベント入力':{'B3':'未完了'}},invalid=True)
    yield Case('fraction_annual',p,edits={'収支':{'D2':0.1}},invalid=True)
    standard=replace(p,end=(2055,12),rules=[replace(r,end=(2055,12)) for r in p.rules],events=[one(id=f'E{i+1:03}',amount=1000,occurrence_month=(2030,1)) for i in range(20)])
    yield Case('performance_standard',standard,36980000,monitor_seconds=30)
    maximum=cash_plan((2026,1),(2125,12))
    maximum.accounts += [Account(f'A{i+1:02}','household','自由資金',0) for i in range(19)]
    maximum.rules=[Rule(f'R{i+1:03}',f'ITEM{i+1:03}','expense',0,(2026,1),(2125,12),'cash') for i in range(50)]
    maximum.events=[one(id=f'E{i+1:03}',amount=100,event_type='periodic',occurrence_month=None,first_month=(2026,1),last_month=(2125,1),interval_months=12) for i in range(200)]
    yield Case('performance_maximum',maximum,-1000000,monitor_seconds=120)


CASES=list(cases())


def prepare(directory,performance=True):
    directory=Path(directory); directory.mkdir(parents=True,exist_ok=True)
    manifest=[]
    for case in CASES:
        if not performance and case.name.startswith('performance_'): continue
        path=directory/f'{case.name}.xlsx'
        w=generate(case.plan,path)
        for sheet,edits in case.edits.items():
            for cell,value in edits.items(): w[sheet][cell]=value
        w.save(path)
        if case.invalid:
            try: calculate(read_plan(path))
            except ValueError: pass
            else: raise AssertionError(f'{case.name}: Pythonが無効入力を拒否しませんでした')
        else:
            normalized=read_plan(path)
            assert calculate(normalized)[-1]['cash']==case.expected_cash,case.name
            w['_検証メタ']['B1']=identity(normalized)[0]
            w.save(path)
        manifest.append(dict(file=path.name,expected_cash=case.expected_cash,invalid=case.invalid,
            monitor_seconds=case.monitor_seconds,sha256=hashlib.sha256(path.read_bytes()).hexdigest()))
        print('prepared',case.name,flush=True)
    (directory/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
    return manifest


if __name__=='__main__':
    parser=argparse.ArgumentParser(); parser.add_argument('--output',default='output/p2a_recalc_input'); parser.add_argument('--without-performance',action='store_true')
    args=parser.parse_args(); prepared=prepare(args.output,not args.without_performance)
    print('prepared cases:',len(prepared))
