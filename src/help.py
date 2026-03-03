<<<<<<< HEAD
S = {
    # generic facility words
    "site","plant","facility","factory","works","complex",
    "terminal","depot","storage","warehouse","hub", "products",
    "unit","phase","project","production","processing","operations","operation", "chem", "chemical", "chemicals", "fuels",

    # transport / port
    "port","harbor","harbour","dock","jetty","berth","quay",
    "tank","tanks","tankfarm","pipeline","loading","unloading","import","export","bunker","bunkering",

    # generic org words
    "group","holding","holdings","international","global","services","solutions","resources","energy", "oil",
    "bio", "global", "renewable", "trading", "industrial", "industria", "industries", "service", "technologies", "company",

    # address words
    "street","st","road","rd","avenue","ave","lane","ln","way","drive","dr","boulevard","blvd","place","pl",
    "building","bldg","block","floor","suite","industrial","estate","zone","park","area",

    # special characters
    "&"
}

exempt_words = ["ltd.", "ltd", "s.i.u",
                    "s.a.", "s.a", "s.r.o.",
                    "s.r.o", "s.i.", "s.i",
                    "s.p.a", "s.p.a.", "s.l.u",
                    "s.l.u", "a.s", "a.s.",
                    "s.l", "s.l.", "inc.", "inc",
                    ". ltd", "-", "oils", "l.p.",
                    "llc", "l.l.c.", "llc.", "lp", "inc.."]

def every_word_has_digit(tok: str) -> bool:
    words = tok.split()
    return bool(words) and all(any(ch.isdigit() for ch in w) for w in words)

cert_owner = "Viterra USA LLC, Warden, WA, United States"



parts = [p.strip().lower() for p in cert_owner.split(",") if p.strip()][1:-1]

# remove company (first) and country (last)

tokens = [
    tok
    for tok in parts
    if tok
    and tok not in exempt_words
    and not any(w in S for w in tok.split())
    and not every_word_has_digit(tok)
]

country = "united states"

# testing to see if this logic works to remove street names coming into the city column by mistake
if len(tokens) == 1:
    print(" ".join([w for w in tokens[0].split() if not w.isnumeric()]).title())
elif len(tokens) >= 2:
    if country in ("united states", "china"):
        print(" ".join([w for w in tokens[-2].split() if not w.isnumeric()]).title())
    else:
        print(" ".join([w for w in tokens[-1].split() if not w.isnumeric()]).title())


=======
t = [1, 2, 3, 4, 5, 6, 7]
print(t[1:-1])
print("hello world".capitalize())
>>>>>>> 619d8a63aa71804771aa4d55458fe215e80faa15
