"""10장 '잔금일 돈 길' 폰 화면 14장의 장면 정의 (n1~n14).

python3 make_scenes.py new scenes.json && node render.js scenes.json out
"""
import json, sys

def r(l, s, st='pend', tm=None, tag=None, **kw):
    d = dict(l=l, s=s, st=st)
    if tm: d['tm'] = tm
    if tag: d['tag'] = tag
    d.update(kw)
    return d

AUTO, BANK, HUMAN = 'auto', 'bank', 'human'

def head(t, n, cap=None, st='이상 없음', red=False, icon=False, prog=None):
    if prog is None:
        prog = ['g'] * n + [''] * (5 - n)
    return dict(t=t, prog=prog, cap=cap or f'조건 5개 중 {n}개 확인', st=st, red=red, icon=icon)

# ---------- corrected ----------
# ③ 법무사 서류 확인 10:00 (사람 확인) → ④ 대출 실행 10:10 (은행 실행)
SPEC = [
    ('등기부 최종 확인', '새 근저당·압류 없음', '새 근저당·압류 없음', '08:30', ['자동 확인', AUTO]),
    ('보증금 2억 원 입금', '지금 집주인 → 잔금 안심금고', '지금 집주인 → 잔금 안심금고', '09:00', ['자동 확인', AUTO]),
    ('법무사 서류 확인', '매도인 등기서류 확인', '매도인 등기서류 확인 완료', '10:00', ['사람 확인', HUMAN]),
    ('대출 2억 4,000만 원 실행', '서류 확인 뒤 실행 · 기존 대출 상환', '매도인 기존 대출은 은행이 바로 상환', '10:10', ['은행 실행', BANK]),
    ('받는 계좌 확인', '예금주 = 등기부상 소유자', '예금주 = 등기부상 소유자', '10:20', ['자동 확인', AUTO]),
]

def new_rows(n, last='lock', last_tm=None):
    rows = []
    for k, (l, s_pend, s_ok, tm, tag) in enumerate(SPEC):
        rows.append(r(l, s_ok, 'ok', tm, tag) if k < n else r(l, s_pend))
    if last == 'lock':
        rows.append(r('잔금 보내기', None, 'pend', rt='잠김'))
    elif last == 'wait':
        rows.append(r('잔금 보내기', None, 'act', rt='승인 대기', rtc='blue'))
    elif last == 'done':
        rows.append(r('잔금 보내기', None, 'ok', tm=last_tm))
    return rows

POST = [
    r('취득세 납부 · 이전등기 접수', '법무사가 취득세를 낸 뒤 바로 접수', 'ok', '11:00', ['접수 확인', 'fin'], hl=True),
    r('전입신고 알림', '이사 후 14일 안에', 'pend', rt='오후'),
    r('대출·감면 기한 알림', '디딤돌 1개월 안 전입 · 감면 3년 유지', 'pend', rt='예약'),
]
POST_RW = [dict(p) for p in POST]
POST_RW[1] = r('전입신고 알림', '이사 후 14일 안에', 'pend', tm='오후', tmc='gray', tag=['알림 예약됨', AUTO])
POST_RW[2] = r('대출·감면 기한 알림', '디딤돌 1개월 안 전입 · 감면 3년 유지', 'pend', tm='예약', tmc='gray', tag=['알림 예약됨', AUTO])

ALERT = ('<div class="alert"><h4>무엇이 어긋났나요?</h4>'
         '<p>지금 집주인의 보증금 2억 원이 09:40까지 들어오지 않았어요.<br>서류 확인·대출 실행과 잔금 송금은 모두 대기 중이에요.</p>'
         '<div class="b1">{msg}<span>중개사님께 보낼 문자 만들기</span></div>'
         '<div class="b2row"><div class="b2">{home}<span>집주인 반환 확인</span></div>'
         '<div class="b2">{phone}<span>법무사·은행 상담</span></div></div></div>')

STOP_ROWS = [
    r('등기부 최종 확인', '새 근저당·압류 없음', 'ok', '08:30'),
    r('보증금 2억 원 입금', '09:40 아직 입금 전', 'bad', tag=['확인 필요', 'need'], sc='red'),
    r('법무사 서류 확인', '매도인 등기서류', 'pend', tag=['대기', 'wait']),
    r('대출 2억 4,000만 원 실행', '보증금·서류 확인 뒤 진행', 'pend', tag=['대기', 'wait']),
    r('받는 계좌 확인', '예금주 = 등기부상 소유자', 'pend', tag=['대기', 'wait']),
    r('잔금 보내기', '2억 1,000만 원', 'pend', tag=['잠김', 'wait']),
]
STOP_HEAD = head('모든 이체를 잠시 멈췄어요', 0, cap='조건 2에서 멈춤', st='확인 필요', red=True, icon=True,
                 prog=['g', 'r', '', '', ''])

ALL5 = head('다섯 가지가 모두 확인됐어요', 5, cap='조건 5개 모두 확인')
NEW = {
    'n1': dict(time='08:25', head=head('오늘은 잔금일이에요', 0), vault=dict(a='1,000만'),
               rows=new_rows(0), send=dict(state='locked')),
    'n2': dict(time='08:30', head=head('1단계까지 확인됐어요', 1), vault=dict(a='1,000만'),
               rows=new_rows(1), send=dict(state='locked')),
    'n3': dict(time='09:00', head=head('2단계까지 확인됐어요', 2), vault=dict(a='2억 1,000만', s='보증금 2억 원이 들어왔어요'),
               rows=new_rows(2), send=dict(state='locked')),
    'n4': dict(time='10:00', head=head('3단계까지 확인됐어요', 3), vault=dict(a='2억 1,000만'),
               rows=new_rows(3), send=dict(state='locked'),
               notif=dict(n='김OO 법무사', m='매도인 등기서류 확인을 완료했어요')),
    'n5': dict(time='10:10', head=head('4단계까지 확인됐어요', 4), vault=dict(a='2억 1,000만'),
               rows=new_rows(4), send=dict(state='locked')),
    'n6': dict(time='10:20', scroll=18, head=ALL5, vault=dict(a='2억 1,000만'),
               rows=new_rows(5, 'wait'), send=dict(state='open')),
    'n7': dict(time='10:30', scroll=18, head=ALL5, vault=dict(a='2억 1,000만'),
               rows=new_rows(5, 'wait'), send=dict(state='open', glow=True)),
    'n8': dict(time='10:30', scroll=18, head=ALL5, vault=dict(a='2억 1,000만'),
               rows=new_rows(5, 'wait'), send=dict(state='open'), overlay='auth'),
    'n9': dict(time='10:31', done=dict(h='2억 1,000만 원을 보냈어요', p1='박OO (매도인) · 우리 ···7310',
               p2='10:31 · 잔금 안심금고에서 출금',
               next='11:00 취득세 납부 · 이전등기 접수 확인<br>→ 전입신고 알림 (14일 안)<br>→ 대출·감면 기한 알림')),
    'n10': dict(time='11:00', scroll=240, head=ALL5, vault=dict(a='0원'),
                rows=new_rows(5, 'done', '10:31'), send=dict(state='sent'), post=POST,
                foot='조건이 하나라도 어긋나면 모든 이체가 멈춰요'),
    'n11': dict(time='09:40', scroll=240, head=ALL5, vault=dict(a='0원'),
                rows=new_rows(5, 'done', '10:31'), send=dict(state='sent'), post=POST_RW,
                foot='조건이 하나라도 어긋나면 모든 이체가 멈춰요', overlay='rewind'),
    'n12': dict(time='09:40', head=STOP_HEAD, alert='new', rows=STOP_ROWS),
    'n13': dict(time='09:40', head=STOP_HEAD, alert='new', rows=STOP_ROWS, overlay='sheet', cursor=True),
    'n14': dict(time='09:40', head=STOP_HEAD, alert='new', rows=STOP_ROWS, overlay='sheet'),
}

if __name__ == '__main__':
    which = sys.argv[1]
    data = NEW
    json.dump(data, open(sys.argv[2], 'w'), ensure_ascii=False)
