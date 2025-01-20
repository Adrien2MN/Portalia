from fastapi import FastAPI

app = FastAPI()

@app.get("/")
def read_root():
    return {"message": "Bienvenue sur FastAPI"}

@app.get("/convert")
def convert(
    tjm: float = None, brut: float = None, net: float = None, 
    jours: int = 18, frais_fixes: float = 0.08, provisions: float = 0.10, 
    charges_sal: float = 0.22, charges_pat: float = 0.12
):
    if tjm:
        brut = 198 + (tjm * jours * (1 - frais_fixes - provisions))
        net = brut * (1 - charges_sal - charges_pat)
    elif brut:
        tjm = (brut - 198) / (jours * (1 - frais_fixes - provisions))
        net = brut * (1 - charges_sal - charges_pat)
    elif net:
        brut = net / (1 - charges_sal - charges_pat)
        tjm = (brut - 198) / (jours * (1 - frais_fixes - provisions))
    return {
    "tjm": round(tjm, 2) if tjm else None,
    "brut": round(brut, 2) if brut else None,
    "net": round(net, 2) if net else None
}
