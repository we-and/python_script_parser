text = """Surprise!
Oh my God, you're huge.
No traffic.
And I crushed the final.

No such luck.
My mom was the math genius.
Look at you.
Wow. Bag packed.
House looks great.
You've been busy.
Where is Miranda?
Yes.
Thank you, thank you.
You sure you're okay?
Seriously, I am here to help.
Mm. Yum.
True.
That's definitely in the genes.
He has that because of you.
Really? You think he's funny?
No wonder he loves you.
Whatever you say.
Oh. I got this for Anthony. Isn't he cute?
And I got this
from my friend who's a psych major.
I think it might help. Dad said Miranda's
still not doing well.
Good.
They should be home soon, right?
Oh.
What is it? What's wrong?
No.
I'll go.
Yes, it makes the most sense.
You should stay here and wait for Miranda to come home.
I'll follow you.
I will.
We'll find her.
I haven't seen him in months.
I've been away at school.
He was so happy with Jenny and the baby.
Still no word on Miranda?
Morning?
That's too late.
I have to do something. I can't just sit here.
I can get volunteers through social media.
Wait, what?
Where's the sheriff? Why isn't he here?
Who's in charge?
But you're not.
My dad is dead, and my sister is missing.
What are you doing about it?
No, you don't.
Jesus.
You called the others, right?
I know, but I just can't sit here anymore
doing nothing.
The deputies aren't telling me anything
and I don't even think they know what they're doing.
I know it's dangerous, you don't have to do it.
Okay. Thank you.
I'll see you in an hour at the rest stop.
And keep it quiet.
Watch your step.
Can we hurry up, please?
 Yeah, I do. He's old.
We should branch out.
Guys! Come here.
That's Miranda's.
Positive.
Does that mean you found her?
Is there something we can do?
Can we go with you
and organize the volunteer search party there?
Please?
Understood.
We wanna help look for Miranda.
Miranda!
Thank God. Thank God.
Oh my God.
I can't get through. It keeps going straight to voicemail.
Hello?
I'll make you a sandwich.
Oh my God.
Oh, my God.
Let's go.
The bag. The bag.
Breathe, Mom. Breathe.
Good.
Oh, my God.

Oh, my girls.
He's coming.
Yeah, we have to go."""

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

with open("count_detail_emily2.csv", 'w', encoding='utf-8') as file:
    file.write(csv)
    print(f"CSV text has been saved")


