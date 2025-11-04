from excels_dup import dupplicate
from reading_moq import read_all_sites, combine_l1_ethpm_from_all, combine_eqpt_top_from_all
from rename import rename_txt
from filling_out_sheets import fill_ports_sheet

if __name__ == '__main__':
    rename_txt()
    all_data = read_all_sites()
    combine_l1_ethpm_from_all(all_data)
    combine_eqpt_top_from_all(all_data)

    """for site, data in all_data.items():
        if "l1PhysIf_ethpmPhysIf" in data and data["l1PhysIf_ethpmPhysIf"]:
            first_obj = data["l1PhysIf_ethpmPhysIf"][0]
            print(f"🏭 {site} → {len(data['l1PhysIf_ethpmPhysIf'])} objetos")
            for k, v in first_obj.items():
                print(f"   {k}: {v}")
            print()
            break
        else:
            print(f"⚠️ {site} → No se encontraron objetos combinados\n")"""
    dupplicate()
    fill_ports_sheet(all_data)