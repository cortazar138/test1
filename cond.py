file = open("test.txt", "r")
lines = file.readlines()
file.close
file = open("test.txt", "w")
k = 0
for line in lines:
    print(lines[1])
    if line == "\n" or line == lines[1]:
        pass
    else:
        file.write(line)

print("koniec")

#for line in ok:
#    file.write(line + "\n")
