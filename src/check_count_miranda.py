text = """No.
Cool.
Yeah, maybe.
Why do you think I'm a vegetarian?
[both chuckle]
Hey, can we stop for a second?
I have to go to the bathroom.
It's-- it's not that.
Fuck off.
He's being weird.
What is he--?
Eddie!
Why is he following, Eddie?
Oh, no! Oh my God.
Eddie!
Eddie?
Oh. Oh my God, Eddie.
Oh my God, we gotta go.
Please.
No, no, no.
Hello?
Hello?
My name is  Cooper.
I'm lost in the middle of the woods.
Please, I need help. My stepdad has been shot.
We just left a rest stop on the main highway.
Please help him.
Hello?
Hello?
Damn it.
Mom, it's me. Um--
Mom, Eddie's been shot. I'm lost in the woods.
Some guy shot him and now he's after me.
And I think I lost him, but I don't know.
Mom, what should I do?
I tried calling 911,
but I don't know what else to do.
Mom.
Shit.
Five, four, three, two, one.
Five, four, three.
Five things I see.
I see a tree.
I see some rocks.
I see the sky.
I see a creek.
Five things I see.
I see a tree.
The river.
A pine cone.
A beetle.
Oh, shit.
Four things I touch.
Tree trunk.
My arms.
A plant.
A leaf.
Three things I hear.
An owl.
A twig.
Oh my God.
Oh my God. He's coming.
He's coming. Please, I need help.
He's-he's old and bald,
and wearing a jean jacket and he has a gun.
Please, I need help. Please.
Oh my God. I'm sorry. There's a crazy man after me.
We have to run, let's go.
There's a man, he shot my stepdad.
We have to run, let's go!
Help! Help!
He's dead, my stepdad.
A guy, he killed my stepdad,
and now he's after me.
No time for that. He has a gun!
You have service?
How can I get service?
What's the password?
No.
Run!
Mom, Eddie's been shot. I'm lost in the woods.
Some guy shot him and now he's after me,
and I think I lost him, but I don't know.
 Mom, what should I do?
I tried calling 911,
but I don't know what else to do. Mom.
Mom?
Yes! Fucking A!
Mom?
I'm okay. Eddie's--
It all happened so fast.
I didn't know what to do.
He told me to run. I didn't want to,
but he kept telling me to go.
I'm so sorry, Mom.
My battery is dying.
Mom?
Mom?
I love you too.
Oh my God. You're alive.
You're hurt?
Are you sure?
Are you okay?
I took it from that man.
I'm sorry, I didn't think.
I was just gonna get it to use the flashlight.
I haven't been using it, I was afraid he'd see the light.
Miranda.
It's still bleeding.
I don't know. He...
followed us from a rest stop and just started shooting.
He killed my stepdad, and I just ran.
There's nothing you could've done.
You're lucky you got away.
It's okay. I'm not hungry.
I-- I called 911 and my mom.
She said people are looking for me.
But he's still out there.
What do you think we should do?
Yeah, I took it from him too.
Sorry, I always shake like that.
Yeah.
How close is it?
Maybe we should make a run for it.
I've never been camping before.
Are you okay?
Here, I have meds. Do you want some?
I'm just supposed to take these every day, but I don't.
Sometimes I have panic attacks.
I haven't taken them in weeks.
Doctor says they'll help.
A year ago, my father killed himself.
And I was the one who found him.
They tried all different kinds of drugs.
They all make me feel like shit.
Please. It might help with the pain.
Me too.
Here.
Mushrooms?
Todd.
Hey Todd, I found some more food.
The sun's rising. We should head out.
Todd.
Shit.
Daddy?
Daddy?
What?
I'm so sorry. I didn't know what to do.
Five things I see.
One. 
One.
Two.
Three.
Four.
Five.
Six.
Seven.
Eight.
Nine.
Help! Help! Help!
Help! He's right behind me!
Emily? 
Here it is. This is where Todd is.
It's almost straight in from here.
I left a backpack to mark the spot.
It was his mom's.
So, that's his name.
Yeah, that's it. And he was pissed.
I'm just so hungry.
Mom?
Get away from her!
Did she do that to your face?
Is that why you chose me?
Is that why you killed my stepdad
and followed me through the woods?
Because I look like your dead wife?
No, Mom.
You're pathetic.
You're nothing.
Eddie and Todd and even my dad,
who was out of his mind, were better men than you.
You're a worthless piece of shit,
and I look like your dead wife?
But I'm not dead. I'm alive!
Ah! Emily! Help me! Help me!
Mom! No, no!
Get off!
Stop!
Mom.
I love you.
I fixed it.
Eddie... Dad...
The last thing he said was...
"Tell your mother I love her."
I'm so sorry, Mom.
What?
Right-- right now?
Oh-- Oh my gosh.
Oh, my God.
Okay. Okay. Okay, Mom. Breathe.
Breathe, Mom. Breathe.
Oh, my God. The bag.
I got the bag.
Come on. I got the bag. Go. Go, go, go.
Daddy! 


"""

# Split the text into lines
lines = text.split('\n')

# Iterate through each line and count the characters
character_counts = {line: len(line.strip()) for line in lines}

# Print the results
charactercount=0
linecount=0
for line, count in character_counts.items():
    print(f"'{line.strip()}' ->  {count} caractères")
    charactercount=charactercount+count
    linecount=linecount+1

print("-----------------------------")
print("Caractères : "+str(charactercount))
print("Lignes     : "+str(linecount))
print("Repliques  : ",str(charactercount/40))


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
print("Lignes     : "+str(linecount))
print("")
print("TOUT")
print("Caractères : "+str(total_charactercount_nopunc))
print("Repliques  : ",str(total_charactercount/40))
print("")
print("PAS DE PONCTUATION")
print("Caractères : "+str(total_charactercount_nopunc))
print("Repliques  : ",str(total_charactercount_nopunc/40))
print("")
print("PAS DE PONCTUATION, PAS D'APOSTROPHE")
print("Caractères : "+str(total_charactercount_nopunc_noap))
print("Repliques  : ",str(total_charactercount_nopunc_noap/40))
print("")
print("PAS DE PONCTUATION, PAS D'ESPACE")
print("Caractères : "+str(total_charactercount_nopunc_nosp))
print("Repliques  : ",str(total_charactercount_nopunc_nosp/40))
print("")
print("PAS DE PONCTUATION, PAS D'APOSTROPHE, PAS D'ESPACE")
print("Caractères : "+str(total_charactercount_nopunc_noap_nosp))
print("Repliques  : ",str(total_charactercount_nopunc_noap_nosp/40))
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
    csv_line=line.strip()+"@"+str(c)+"@"+str(c_nopunc)+"@"+str(c_nopunc_noap)+"@"+str(c_nopunc_nosp)+"@"+str(c_nopunc_noap_nosp)
    #print(csv_line)
    csv=csv+csv_line+"\n"
    i=i+1

with open("count_detail_miranda.csv", 'w', encoding='utf-8') as file:
    file.write(csv)
    print(f"CSV text has been saved")


