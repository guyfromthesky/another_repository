malformed_name = input("Enter the malformed file name:")
corrected_name = malformed_name.encode('cp437').decode('euc_kr')
print('Corrected file name: ', corrected_name)