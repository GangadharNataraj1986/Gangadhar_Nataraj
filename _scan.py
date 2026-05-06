lines = open('Problem_Solution_Agent_PSS/AI_Assisted_PSS.py', encoding='utf-8').readlines()
for i in range(2551, 2882):
    if '"""' in lines[i]:
        print(f'{i+1:4}: {repr(lines[i][:80])}')
print('done')
