text = """Nice to get away, huh?
Yeah. Remote.
Remote, huh?
Everything okay in there?
He bothering you?
Boy, you're a real mess.
Look at that.
Brought a knife to a gunfight, didn't even use it.
Mm.
Spicy.
What's your code?
So peaceful here, isn't it?
Remote.
Thank you.
Aw.
That's a good one.
Jenny.
God damn.
You're a lucky man.
Were.
Jesus Christ.
Come on.
Need help?
Like this?
Miranda?
Miranda?
Miranda?
Come on.
Stop being a little bitch.
Come on out.
I'm not gonna hurt you.
Come on.
Be a good girl now. Come on out.
I'll be your new daddy.
Miranda, don't make me angry.
Oh.
It's delicious. Thank you, hon.
Really?
Nance.
Excuse me.
Oh.
Hi, Cupcake.
Hi, Cupcake.
You better stop calling. Eddie might get jealous.
I can do this all night.
It's fine.
She's not alone.
What the--?
No, no, no, no. No.
Fuck.
Fuck.
Nice fucking trick.
Doesn't matter. You're all the same.
Fucking deceptive, manipulative fucking cunts.
She's not as smart as she thinks she is.
Well, she's about to be a dead little girl.
Survival,
the act or fact of living
or continuing longer than another person or thing.
That's the Merriam-Webster dictionary definition.
So basically, I don't have to run faster
than the man-eating tiger,
I just have to run faster than you.
That's what I mean by strategic default.
Say the value of your house
drops below what you owe on the mortgage.
What do you do?
What do you do?
You stop making payments.
Now it's not your problem. It's the bank's problem.
And fuck the bank.
So, if your business is going under, you just--
tuck your tail between your legs?
Close up shop? No.
No.
You take as much unsecured debt as you can,
in the name of the business, of course.
You cut expenses to the bone.
You leverage that cash to open a new business.
By the time the bank figures out
your old business is dead, you're on to your next target.
The key is to stay in the game.
Keep moving before anybody can figure out
what you're doing.
You keep moving forward like a shark. Right?
You're either predator or you're prey.
Survival of the fittest.
A simple concept,
but it's not as easy as it sounds.
You gotta take risks.
Why do I waste my time?
You know what, cut the camera.
Everybody get out.
What? I'm letting you go early, okay?
You're welcome.
Go. Everybody get the hell out of my classroom.
Robert.
Make sure and post the lecture today.
Don't fuck it up like last time.
And cancel my office hours. I'm going home.
Miranda?
Miranda?
Miranda?
Hi, Cupcake.
Calm down, Jenny.
I want you to kill me.
Maybe I am. Maybe I am.
Take the fucking gun and shoot me.
Please.
Here, take the gun.
Stop asking the same question over and over, Jenny.
Here I am. Do it!
I did.
She's dead, Jenny! I fucking killed her!
Don't you get that?
Come on! Come on. Come here. Come here.
Come here. You can do it.
We'll do it together, okay? Come on.
I know you can do it.
You have to rack the gun, dummy.
Hi, Emily Cooper?
My name is Randy. I'm a nurse here
at the maternity ward,
and uh-- your mom, or your stepmom, I guess,
Jenny, is here with us.
She's very concerned about Miranda.
Can you tell me about her?
Uh-huh. Oh, that's great news.
She'll be very relieved.
Yeah. Yes, she is in labor,
but I'm sure it's still several hours away.
She would love it if you could do her a big favor
and stop by the house
to pick up her overnight bag that she left behind.
Yes. Yes, very exciting.
Little brother's on the way. Okay, bye-bye.
Our lucky day. They're on the way.
Welcome.
Hey, hey, hey.
Have a seat.
Wait.
Take it off.
Miranda, you're here next to me.
Hello, Emily Cooper. Nice to meet you in person.
We spoke on the phone.
Thanks for bringing Miranda home.
Have a seat.
This is nice.
Let's hold hands.
Bow your heads.
Close your eyes.
Dear Lord,
thank you for your random acts of betrayal,
adultery, wrath,
and violence,
for destroying families
and lives,
leaving nothing but death
and destruction in your wake.
Thank you for the rage and the thrill of the chase.
And thank you for these three beautiful women...
and for the glorious day ahead.
Amen.
Okay, the hand-holding is over now.
Let go.
Let's get started, shall we?
Sometimes when my wife was in the mood,
she enjoyed...
sharing a bottle of bubbly.
Ah. My wife.
So beautiful.
Now?
She's at home.
Where I left her.
In bed.
Dead.
You like it?
God, you look so much like her.
You're a quick learner, Cupcake.
It doesn't matter."""

# Split the text into lines
lines = text.split('\n')

# Iterate through each line and count the characters
character_counts = {line: len(line.strip()) for line in lines}
character_counts_nopunc = {line: len(line.strip().replace(".","").replace("?","").replace("!","")) for line in lines}
character_counts_nopunc_noap= {line: len(line.strip().replace(".","").replace("?","").replace("!","").replace("'","")) for line in lines}
character_counts_nopunc_noap_nosp= {line: len(line.strip().replace(".","").replace("?","").replace("!","").replace("'","").replace(" ","")) for line in lines}
character_counts_nopunc_nosp= {line: len(line.strip().replace(".","").replace("?","").replace("!","").replace(" ","")) for line in lines}

character_count_arr=[]
character_count_nopunc_arr=[]
character_count_nopunc_noap_arr=[]
character_count_nopunc_noap_nosp_arr=[]
character_count_nopunc_nosp_arr=[]

# Print the results
linecount=0
for line, count in character_counts.items():
    linecount=linecount+1


total_charactercount=0
i=1
for line, count in character_counts.items():
    print(f"'{line.strip()}' ->  {count} caractères")
    total_charactercount=total_charactercount+count    
    character_count_arr.append(count)
    i=i+1


total_charactercount_nopunc=0
i=1
for line, count in character_counts_nopunc.items():
    print(f"'{line.strip()}' ->  {count} caractères")
    total_charactercount_nopunc=total_charactercount_nopunc+count    
    character_count_nopunc_arr.append(count)
    i=i+1



total_charactercount_nopunc_noap=0
i=1
for line, count in character_counts_nopunc_noap.items():
    print(f"'{line.strip()}' ->  {count} caractères")
    total_charactercount_nopunc_noap=total_charactercount_nopunc_noap+count    
    character_count_nopunc_noap_arr.append(count)
    i=i+1

total_charactercount_nopunc_nosp=0
i=1
for line, count in character_counts_nopunc_nosp.items():
    print(f"'{line.strip()}' ->  {count} caractères")
    total_charactercount_nopunc_nosp=total_charactercount_nopunc_nosp+count    
    character_count_nopunc_nosp_arr.append(count)
    i=i+1


total_charactercount_nopunc_noap_nosp=0
i=1
for line, count in character_counts_nopunc_noap_nosp.items():
    print(f"'{line.strip()}' ->  {count} caractères")
    total_charactercount_nopunc_noap_nosp=total_charactercount_nopunc_noap_nosp+count    
    character_count_nopunc_noap_nosp_arr.append(count)
    i=i+1




print("-----------------------------")
print("Lignes              : "+str(linecount))
print("")
print("TOUT")
print("Caractères          : "+str(total_charactercount_nopunc))
print("Repliques           : ",str(total_charactercount/40))
print("")
print("PAS DE PONCTUATION")
print("Caractères          : "+str(total_charactercount_nopunc))
print("Repliques           : ",str(total_charactercount_nopunc/40))
print("")
print("PAS DE PONCTUATION, PAS D'APOSTROPHE")
print("Caractères          : "+str(total_charactercount_nopunc_noap))
print("Repliques           : ",str(total_charactercount_nopunc_noap/40))
print("")
print("PAS DE PONCTUATION, PAS D'ESPACE")
print("Caractères          : "+str(total_charactercount_nopunc_nosp))
print("Repliques           : ",str(total_charactercount_nopunc_nosp/40))
print("")
print("PAS DE PONCTUATION, PAS D'APOSTROPHE, PAS D'ESPACE")
print("Caractères          : "+str(total_charactercount_nopunc_noap_nosp))
print("Repliques           : ",str(total_charactercount_nopunc_noap_nosp/40))
print("")
print(len(character_count_arr))
print(len(character_count_nopunc_arr))
print(len(character_counts.items()))
csv=""
i=0
for line,count in character_counts.items():
    c=character_count_arr[i]
    c_nopunc=character_count_nopunc_arr[i]
    c_nopunc_noap_nosp=character_count_nopunc_noap_nosp_arr[i]
    c_nopunc_noap=character_count_nopunc_noap_arr[i]
    c_nopunc_nosp=character_count_nopunc_nosp_arr[i]
    csv_line=line+"@"+str(c)+"@"+str(c_nopunc)+"@"+str(c_nopunc_noap)+"@"+str(c_nopunc_nosp)+"@"+str(c_nopunc_noap_nosp)
    #print(csv_line)
    csv=csv+csv_line+"\n"
    i=i+1

with open("count_detail_wade.csv", 'w', encoding='utf-8') as file:
    file.write(csv)
    print(f"CSV text has been saved")


