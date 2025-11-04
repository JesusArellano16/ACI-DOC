import os
import re
from collections import defaultdict

# --- Definir rutas base ---
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # IMPRESOS/
SRC_DIR = os.path.join(BASE_DIR, "src")

# --- Regex para detectar moquery y líneas de datos ---
re_moquery = re.compile(r'moquery\s+-c\s+(\S+)', re.IGNORECASE)
re_key_value = re.compile(r'^\s*([\w#.]+)\s*:\s*(.*)$')

def parse_txt_file(txt_path):
    site_name = os.path.splitext(os.path.basename(txt_path))[0][:17]

    with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
        lines = [line.rstrip() for line in f if line.strip()]

    site_data = defaultdict(list)
    current_class = None
    current_obj = {}

    for line in lines:
        mo_match = re_moquery.search(line)
        if mo_match:
            if current_class and current_obj:
                site_data[current_class].append(current_obj)
                current_obj = {}
            current_class = mo_match.group(1)
            continue

        if line.strip().startswith("# "):
            if current_class and current_obj:
                site_data[current_class].append(current_obj)
            current_obj = {"#": line.strip()}
            continue

        kv = re_key_value.match(line)
        if kv and current_class:
            key, value = kv.groups()
            key = key.strip()
            value = value.strip()

            # --- Solo para moquery l3extRsPathL3OutAtt ---
            if current_class.lower() == "l3extrspathl3outatt" and key == "dn":
                if not value.endswith("]]"):
                    # Agregar sufijo faltante
                    value = value + "L3_RMAC1]]"

            current_obj[key] = value

    if current_class and current_obj:
        site_data[current_class].append(current_obj)

    return site_name, dict(site_data)


def read_all_sites():
    all_data = {}

    for txt_file in os.listdir(SRC_DIR):
        if txt_file.lower().endswith(".txt"):
            txt_path = os.path.join(SRC_DIR, txt_file)
            site, parsed_data = parse_txt_file(txt_path)
            all_data[site] = parsed_data

    return all_data

def parse_dn(dn):
    """Extrae pod, node y phys del DN."""
    pod_match = re.search(r'pod-(\d+)', dn)
    node_match = re.search(r'node-(\d+)', dn)
    phys_match = re.search(r'phys-\[([^\]]+)\]', dn)

    return {
        'pod': pod_match.group(1) if pod_match else None,
        'node': node_match.group(1) if node_match else None,
        'phys': phys_match.group(1) if phys_match else None
    }


def combine_l1_ethpm_from_all(all_data):
    """Combina l1PhysIf y ethpmPhysIf para todos los sitios."""
    for site, moqueries in all_data.items():
        if "l1PhysIf" not in moqueries or "ethpmPhysIf" not in moqueries:
            continue

        l1_list = moqueries["l1PhysIf"]
        ethpm_list = moqueries["ethpmPhysIf"]
        combined = []

        for l1 in l1_list:
            dn_l1 = l1.get('dn', '')
            if not dn_l1:
                continue
            match_eth = next(
                (e for e in ethpm_list if e.get('dn', '').startswith(dn_l1)),
                None
            )

            parsed = parse_dn(dn_l1)
            combo = {
                'adminSt': l1.get('adminSt'),
                'descr': l1.get('descr'),
                'dn_1': dn_l1,
                'pod': parsed['pod'],
                'node': parsed['node'],
                'phys': parsed['phys']
            }

            if match_eth:
                combo.update({
                    'dn_2': match_eth.get('dn'),
                    'operSt': match_eth.get('operSt'),
                    'operStQual': match_eth.get('operStQual')
                })
            combined.append(combo)
        moqueries["l1PhysIf_ethpmPhysIf"] = combined

def combine_eqpt_top_from_all(all_data):
    """Combina eqptCh y topSystem para todos los sitios usando el dn (topology/pod-X/node-Y)."""
    for site, moqueries in all_data.items():
        if "eqptCh" not in moqueries or "topSystem" not in moqueries:
            continue

        eqpt_list = moqueries["eqptCh"]
        topsys_list = moqueries["topSystem"]
        combined = []

        for eqpt in eqpt_list:
            dn_eqpt = eqpt.get('dn', '')
            if not dn_eqpt:
                continue

            # Extraer pod y node del dn
            pod_match = re.search(r'pod-(\d+)', dn_eqpt)
            node_match = re.search(r'node-(\d+)', dn_eqpt)
            pod = pod_match.group(1) if pod_match else ''
            node = node_match.group(1) if node_match else ''

            # Buscar topSystem cuyo dn coincida con el mismo pod y node
            match_top = next(
                (t for t in topsys_list
                 if f"pod-{pod}/node-{node}/" in t.get('dn', '')),
                None
            )

            combo = {
                'dn_eqpt': dn_eqpt,
                'model': eqpt.get('model'),
                'pod': pod,
                'node': node,
            }

            if match_top:
                combo.update({
                    'dn_top': match_top.get('dn'),
                    'address': match_top.get('address'),
                    'name': match_top.get('name'),
                    'serial': match_top.get('serial'),
                    'version': match_top.get('version')
                })

            combined.append(combo)

        # Guardar resultado combinado
        moqueries["eqptCh_topSystem"] = combined



if __name__ == "__main__":
    all_data = read_all_sites()
    combine_l1_ethpm_from_all(all_data)
    combine_eqpt_top_from_all(all_data)

    # Mostrar ejemplo del primer sitio
    for site, data in all_data.items():
        if "l1PhysIf_ethpmPhysIf" in data:
            print(f"\n🧩 {site} → {len(data['l1PhysIf_ethpmPhysIf'])} objetos combinados\n")
            print(data["l1PhysIf_ethpmPhysIf"][0])
            break