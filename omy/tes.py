name = "Booked by "
name11 = "Booked by"
name1="Status "
name2 = "guest"
a = name11.split(" ")
a1 = name.split(" ")
print(a)
print(a1)
if a[-1] == "":
    b = " ".join(a[:-1])
else:
    b = "".join(a)
c = b +"."
print(c)