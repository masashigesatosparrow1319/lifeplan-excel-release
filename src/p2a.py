"""P2-A extension; the accepted P1 implementation is imported unchanged."""
from dataclasses import dataclass, field, replace
from datetime import date
from collections import defaultdict
import re
import hashlib

from .lifeplan import Plan, Rule, Person, Account, allocate, bounded_integer, months, MAX_AMOUNT

MAX_MONTHS, MAX_RULES, MAX_EVENTS, MAX_ACCOUNTS = 1200, 50, 200, 20
CODE = re.compile(r'[A-Za-z0-9_-]{1,64}\Z')


def month_number(ym):
    if not isinstance(ym, (tuple, list)) or len(ym) != 2:
        raise ValueError('年月は年/月の組です')
    y, m = ym
    if type(y) is not int or type(m) is not int or not 1900 <= y <= 9999 or not 1 <= m <= 12:
        raise ValueError('年月は1900〜9999年、1〜12月です')
    return 12*y+m-1


def code(value):
    if not isinstance(value, str) or not CODE.fullmatch(value):
        raise ValueError('コードは英数字・_・-の1〜64文字です')
    return value.upper()


def lookup(value):
    # Legacy P1 reference IDs may be Japanese; only NEW machine cost codes
    # and event IDs need the restricted alphabet.
    if not isinstance(value,str) or not value: raise ValueError('参照IDが必要です')
    return value.upper()


@dataclass
class CashEvent:
    id: str
    name: str
    kind: str
    amount: int
    account_id: str
    event_type: str = 'once'
    occurrence_month: tuple | None = None
    first_month: tuple | None = None
    interval_months: int | None = None
    last_month: tuple | None = None
    person_id: str = 'household'
    purpose: str = 'その他'
    economic_key: str = ''
    additional_confirmed: bool = False
    enabled: bool = True
    inflation_enabled: bool = False
    base_month: tuple | None = None
    rate_bps: int = 0
    linked_dream_id: str = ''
    note: str = ''

    @property
    def key(self):
        return code(self.economic_key or self.id)


@dataclass
class P2Plan(Plan):
    events: list = field(default_factory=list)
    rule_keys: dict = field(default_factory=dict)

    def keys(self):
        # Give legacy non-code targets readable, stable template group codes.
        keys = {}
        for r in self.rules:
            if r.id in self.rule_keys:
                key = code(self.rule_keys[r.id])
            elif isinstance(r.target, str) and CODE.fullmatch(r.target):
                key = code(r.target)
            else:
                key = 'R_GROUP_'+hashlib.sha256(r.target.upper().encode('utf-8')).hexdigest()[:16].upper()
            keys[r.id] = key
        return keys

    def validate(self):
        # P2-A lookup codes are case-insensitive, matching Excel's lookups.
        # Use a defensive P1 validation view; accepted P1 models/code stay intact.
        persons=[]
        for p in self.persons:
            if type(p.planned) not in (bool,int,float) or p.planned not in (0,1): raise ValueError('出生予定は真偽値です')
            persons.append(replace(p,id=lookup(p.id),role=p.role.lower() if isinstance(p.role,str) else p.role,planned=bool(p.planned)))
        accounts=[]
        for a in self.accounts:
            if type(a.restricted) is not bool: raise ValueError('利用制限は真偽値です')
            owner=lookup(a.owner)
            accounts.append(replace(a,id=lookup(a.id),owner='household' if owner=='HOUSEHOLD' else owner))
        rules=[replace(r,id=lookup(r.id),account=lookup(r.account),kind=r.kind.lower() if isinstance(r.kind,str) else r.kind) for r in self.rules]
        Plan(self.start,self.end,persons,accounts,rules).validate()
        start, end = month_number(self.start), month_number(self.end)
        if end-start+1 > MAX_MONTHS or len(self.rules)>MAX_RULES or len(self.events)>MAX_EVENTS or len(self.accounts)>MAX_ACCOUNTS:
            raise ValueError('P2-A容量超過: 1200月/Rule50/イベント200/口座20')
        for r in self.rules:
            lookup(r.id)
            if not isinstance(r.target,str) or not r.target.strip(): raise ValueError('収入/支出対象が必要です')
            month_number(r.start); month_number(r.end)
        grouped=defaultdict(list)
        for r in self.rules: grouped[r.target.upper()].append(r)
        for group in grouped.values():
            ordered=sorted(group,key=lambda r:r.start)
            if any(b.start<=a.end for a,b in zip(ordered,ordered[1:])):
                raise ValueError('同一対象の期間重複')
        ids = {lookup(r.id) for r in self.rules}
        if len(ids) != len(self.rules):
            raise ValueError('財務ソースID重複')
        if len({lookup(a.id) for a in self.accounts}) != len(self.accounts) or len({lookup(p.id) for p in self.persons}) != len(self.persons):
            raise ValueError('口座/人物ID重複')
        if set(self.rule_keys)-{r.id for r in self.rules}: raise ValueError('不明なRuleキー設定')
        self.keys()
        accounts = {lookup(a.id):a for a in self.accounts}
        people = {lookup(p.id) for p in self.persons}|{'HOUSEHOLD'}
        for e in self.events:
            if type(e.enabled) is not bool: raise ValueError('有効は真偽値です')
            if not e.enabled: continue
            identity = code(e.id)
            if identity in ids: raise ValueError('財務ソースID重複: '+e.id)
            ids.add(identity)
            if not isinstance(e.economic_key,str): raise ValueError('費用項目コードは文字列です')
            e.key
            if not isinstance(e.name, str) or not e.name.strip(): raise ValueError('イベント名が必要です')
            bounded_integer(e.amount,MAX_AMOUNT,'イベント金額')
            if e.kind not in ('income','expense') or e.event_type not in ('once','periodic'):
                raise ValueError('イベント区分が不正です')
            if e.additional_confirmed is not True: raise ValueError('追加費用/収入の確認が必要です')
            if lookup(e.account_id) not in accounts or accounts[lookup(e.account_id)].restricted:
                raise ValueError('口座不明または利用制限口座です')
            if lookup(e.person_id) not in people: raise ValueError('対象者不明')
            if not isinstance(e.purpose,str) or not e.purpose.strip(): raise ValueError('用途が必要です')
            if e.event_type=='once':
                month_number(e.occurrence_month)
                if any(x is not None for x in (e.first_month,e.interval_months,e.last_month)):
                    raise ValueError('一時と周期の入力競合')
            else:
                if e.kind!='expense' or e.occurrence_month is not None: raise ValueError('周期は支出だけです')
                first,last=month_number(e.first_month),month_number(e.last_month)
                bounded_integer(e.interval_months,120000,'周期月数')
                if e.interval_months==0 or last<first: raise ValueError('周期設定が不正です')
            if e.inflation_enabled:
                if e.kind!='expense': raise ValueError('物価補正は支出だけです')
                base=month_number(e.base_month)
                first=month_number(e.occurrence_month if e.event_type=='once' else e.first_month)
                if first<base: raise ValueError('基準月より前の補正発生は禁止です')
                if type(e.rate_bps) not in (int,float) or not -1000<=e.rate_bps<=1000 or e.rate_bps!=int(e.rate_bps):
                    raise ValueError('年率は-1000〜1000の整数bpです')
            if type(e.inflation_enabled) is not bool: raise ValueError('補正有効は真偽値です')


def occurs(e, ym):
    k=month_number(ym)
    if not e.enabled: return False
    if e.event_type=='once': return k==month_number(e.occurrence_month)
    first,last=month_number(e.first_month),month_number(e.last_month)
    return first<=k<=last and (k-first)%int(e.interval_months)==0


def scheduled_months(e,start,end):
    lower,upper=month_number(start),month_number(end)
    if not e.enabled: return
    if e.event_type=='once':
        k=month_number(e.occurrence_month)
        if lower<=k<=upper: yield (k//12,k%12+1)
    else:
        first,last=month_number(e.first_month),month_number(e.last_month)
        interval=int(e.interval_months)
        k=first+max(0,(lower-first+interval-1)//interval)*interval
        while k<=min(last,upper):
            yield (k//12,k%12+1)
            k+=interval


def corrected_price(e,ym,cache=None):
    if not e.inflation_enabled: return int(e.amount)
    n=(month_number(ym)-month_number(e.base_month))//12
    if n<0: raise ValueError('基準月より前の補正発生は禁止です')
    prices=cache.setdefault(e.id,[int(e.amount)]) if cache is not None else [int(e.amount)]
    while len(prices)<=n:
        value=(prices[-1]*(10000+int(e.rate_bps))+5000)//10000
        if value>MAX_AMOUNT: raise ValueError('補正価格が10億円を超えます')
        prices.append(value)
    return prices[n]


def calculate(plan):
    plan.validate()
    keys=plan.keys()
    balances={a.id:int(a.balance) for a in plan.accounts}
    account_ids={lookup(a.id):a.id for a in plan.accounts}
    available={a.id for a in plan.accounts if not a.restricted}
    allocations={r.id:allocate(r.annual,r.weights) for r in plan.rules}
    due=defaultdict(list)
    prices={}
    for e in plan.events:
        if e.enabled:
            for ym in scheduled_months(e,plan.start,plan.end):
                due[ym].append((e,corrected_price(e,ym,prices)))
    rows=[]
    for ym in months(plan.start,plan.end):
        opening=sum(balances[a] for a in available)
        ledger=[]
        occupied=set()
        for r in plan.rules:
            if r.start<=ym<=r.end:
                ledger.append(dict(source_id=r.id,key=keys[r.id],kind=r.kind.lower(),amount=allocations[r.id][ym[1]-1],account=account_ids[lookup(r.account)]))
        for e,value in due[ym]:
            ledger.append(dict(source_id=e.id,key=e.key,kind=e.kind,amount=value,account=account_ids[lookup(e.account_id)],
                               dream_id=e.linked_dream_id,purpose=e.purpose,person_id=e.person_id))
        income=expense=0
        for item in ledger:
            value=item['amount']
            if value>0:
                if item['key'] in occupied: raise ValueError(f'二重計上: {ym} {item["key"]}')
                occupied.add(item['key'])
            if item['kind']=='income':
                income+=value; balances[item['account']]+=value
            else:
                expense+=value; balances[item['account']]-=value
        cash=sum(balances[a] for a in available)
        rows.append(dict(month=ym,opening=opening,income=income,expense=expense,cash=cash,total=sum(balances.values()),
                         shortage=max(0,-cash),accounts=dict(balances),ledger=ledger))
    return rows


def from_p1(plan,events=None):
    return P2Plan(plan.start,plan.end,plan.persons,plan.accounts,plan.rules,list(events or []))
