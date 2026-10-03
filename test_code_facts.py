from types import SimpleNamespace as B
def mk(mode):
    def f(start, end, bookings):
        for b in list(bookings) if mode=="remove" else bookings:
            if b.status == "cancelled":
                if mode=="orig": return True
                if mode=="continue": continue
                if mode=="break": break
                if mode=="pass": pass
                if mode=="remove": bookings.remove(b)
            if start < b.end and end > b.start: return False
        return True
    return f
def bk(cancel_last=False):
    c=B(status="cancelled",start=6,end=7); k=B(status="confirmed",start=7,end=8)
    return [k,c] if cancel_last else [c,k]
T=lambda s,e:(s,e)
def chk(mode,s,e,last=False): return mk(mode)(s,e,bk(last))
assert chk("orig",6.5,7.5) is True
assert chk("continue",6.5,7.5) is False and chk("continue",6,7) is True and chk("continue",8,9) is True
assert chk("break",6.5,7.5) is True
assert chk("pass",6,7) is False
assert chk("remove",6,7) is False
assert chk("orig",6.5,7.5,True) is False
assert chk("orig",8,7) is True  # no start<end validation
print("all code facts OK")
