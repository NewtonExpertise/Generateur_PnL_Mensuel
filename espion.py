import getpass
import logging
from collections import OrderedDict
from datetime import datetime
from postgreagent import PostgreAgent


def update_espion(dossier = "", base = "", argStr = ""):

    conf = OrderedDict(
        [
            ('host', '10.0.0.17'), 
            ('user', 'admin'), 
            ('password', 'Zabayo@@'), 
            ('port', '5432'), 
            ('dbname', 'outils')
            ]
        )

    horodat = datetime.now()
    collab = getpass.getuser()
    table = "pnl"
    values = [collab, horodat, dossier, base, "PNL_MENSUEL"]

    sql = f"""
    INSERT INTO pnl (collab, horodat, code_client, base, operation) 
    VALUES ('{collab}', '{horodat}', '{dossier}', '{base}', 'PNL_MENSUEL');
    """
    print(sql)
    with PostgreAgent(conf) as db:
        if db.connection:
            if db.table_exists(table):
                logging.debug(f"table {table} exists")        
                print(db.query(sql))

# if __name__ == "__main__":
#     import logging
#     update_espion("FORM05", "abc")

