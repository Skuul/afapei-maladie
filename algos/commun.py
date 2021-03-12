from collections import defaultdict, namedtuple
from datetime import datetime, timedelta
from typing import List, Dict
import openpyxl
from os import sep as fs_sep

Enregistrement = namedtuple('Enregistrement',
    ['Matricule',
    'Nom',
    'Prenom',
    'Code_service',
    'Service',
    'Date_planning',
    'Motif',
    'Libelle_du_motif',
    'Heures',
    'Arret_initial']
)
Salarie = namedtuple('Salarie', ['matricule', 'nom', 'prenom'])
Absence = namedtuple('Absence', ['debut', 'fin', 'motif', 'heures', 'heures_3_premiers_jours'])

def pasPlusDeTroisJours(d1:datetime, d2:datetime):
    return (d2 - d1).days < 3

def formatAbsence(absence):
    if absence.Motif == 'MAL':
        return Absence(absence.Date_planning, absence.Date_planning, absence.Motif, absence.Heures, absence.Heures)
    else:
        return Absence(absence.Date_planning, absence.Date_planning, absence.Motif, absence.Heures, '')

def filterAbsences(base: List[Enregistrement]):
    d = defaultdict(list)
    for enr in base:
        d[enr.Matricule].append(enr)

    # Rend les données sous forme de générateur
    yield from map(
        lambda v: (
            Salarie(v[0], v[1][0].Nom, v[1][0].Prenom),
            sorted(v[1], key=lambda a: a.Date_planning)
        ),
        d.items()
    )

def estJourSuivant(d1:datetime, d2:datetime):
    modified_date = d1 + timedelta(days=1)
    s1 = datetime.strftime(modified_date, "%Y/%m/%d")
    s2 = datetime.strftime(d2, "%Y/%m/%d")
    return s1 == s2
