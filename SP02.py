#python program to modify SP02 txt output
file = open("test.txt", "r")
lines = file.readlines()
file.close
file = open("test.txt", "w")
for line in lines:
    print(lines[1])
    if line == "\n" or line == lines[1]:
        pass
    else:
        file.write(line)

print("koniec")
