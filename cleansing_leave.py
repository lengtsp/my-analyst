def prepare_leaveadjust(filename):
    import pandas as pd
    df            = pd.read_excel(filename)
    df            = df.rename(columns={df.columns[0] : 'EMPLID'}      )
    df['leave_type_k'] = 0
    df['leave_type_k'] = df['Leave Type'].apply(convert_leavetype)
    return df


def convert_leavetype(ty):
    if ty == 'ANL1':
        return 8
    elif ty == 'ANL2':
        return 8
    elif ty == 'ANLCARR':
        return 34
    elif ty == 'HJL1':
        return 13
    elif ty == 'HJL2':
        return 13
    elif ty == 'MSL':
        return 10
    elif ty == 'MTL_1':
        return 4
    elif ty == 'MTL_2':
        return 33
    elif ty == 'MTL_3':
        return 5

    elif ty == 'MTL_INF1':
        return 4
    elif ty == 'MTL_INF2':
        return 33
    elif ty == 'OLP1.1':
        return 6
    elif ty == 'OLP1.2':
        return 6
    elif ty == 'OLP2':
        return 6

    elif ty == 'PTL':
        return 31
    elif ty == 'PVL':
        return 2
    elif ty == 'SL':
        return 3
    elif ty == 'TNL':
        return 12

    elif ty == 'VPSNL':
        return 30
  
    elif ty == 'STL':
        return 19


    else:
        return 'N/A'

   
