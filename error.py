#for checking the line
def error (kadr, tx):
    try:
        if kadr== 'tarikh':
            if len(tx)!= 10:
                return False
            if type(int(tx[0:4]))== int and tx[4]=='/' and type(int(tx[5:7]))==int and tx[7]=='/' and type(int(tx[8:10]))== int:
                return True
            else:
                return False
        if kadr== 'vazn' or 'tedad':
            b=int(tx)
            return True
    except ValueError:
        return False
        