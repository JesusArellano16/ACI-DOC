from excels_dup import dupplicate
from reading_moq import read_all_sites, combine_l1_ethpm_from_all, combine_eqpt_top_from_all
from rename import rename_txt
from filling_out_sheets import fill_ports_sheet

if __name__ == '__main__':
    rename_txt()
    all_data = read_all_sites()
    combine_l1_ethpm_from_all(all_data)
    combine_eqpt_top_from_all(all_data)
    dupplicate()
    fill_ports_sheet(all_data)