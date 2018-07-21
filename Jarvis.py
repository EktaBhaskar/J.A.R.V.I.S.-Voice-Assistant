import speech_recognition as sr #it recognises speech 
import webbrowser
import wolframalpha #mathematical operations
import wikipedia #for any info from wikipedia
import time
import os #for operating system
import pyperclip #to copy paste
import win32com.client as winc1 #text to speech
v=winc1.Dispatch("SAPI.SpVoice") #way to initialise speech
c1=wolframalpha.Client('4TJERH-2QPVW7EUW3') #to create a object of the app id ,we can use it to call function
att=c1.query('Test/Attempt')
r=sr.Recognizer()
r.pause_threshold=0.7
r.energy_threshold=400

shell=winc1.Dispatch("WScript.Shell") # used to handle interupts of keyboard
v.Speak('Hello!For a list of commands,please say "keyword list"...')
print('Hello!For a list of commands,please say "keyword list"...')
keywd='keyword list'
google='search for'
acad='academic search'
sc='deep search'
wkp='wiki page for'
rdds='read this text'
sav='save this text'
bkmk='bookmark this page'
vid='video for'
wtis='what is'
wtar='what are'
whis='who is'
whws='who was'
when='when'
where='where'
how='how'
paint='open paint'
lsp='silence please'
lsc='resume listening'
stoplst='stop listening'

while True:
	with sr.Microphone() as source: #when user has to speak, input is taken through microphone
		try:
			print("Please Speak")
			audio=r.listen(source, timeout=None)
			message=str(r.recognize_google(audio))
			print('You said:'+message)
		#except:
		#	break
		#finally:
		#	break
			if google in message: #for comparing words in message from keywords
				words=message.split() #spliting words that we speak
				del words[0:2] #delete initial 2 words from the message we speak
				st=' '.join(words) #join those words and make a string 
				print('Google Results for: '+str(st)) #prints the string
				url='http://google.com/search?q='+st #string goes in search (query in q)
				webbrowser.open(url) #opens the webbrowser for the url
				v.speak('Google Results for: '+str(st)) #speaks what we have searched
			elif acad in message: # for searching for string containing acad keyword
				words=message.split()
				del words[0:2]
				st=' '.join(words)
				print('Academic results for: '+str(st))
				url='https://scholar.google.com/scholar?q='+st
				webbrowser.open(url)
				v.speak('Academic results for:'+str(st))
			elif wkp in message: #speak wikipage for
				try:
					words=message.split()
					del words[0:3]
					st=' '.join(words)
					wkpres=wikipedia.summary(st,sentences=2) #take two initial lines of wikipedia summary
					try:
						print('\n'+str(wkpres)+'\n') #prints the line
						v.Speak(wkpres) #speak the lines printed
					except UnicodeEncoderError: #when there is no word in the page or other error
						v.Speak(wkpres) #exception found is spoken,v can also use eception instead of wkpres
				except wikipedia.exceptions.DisambiguationError as e: #when there r ambiguous words like python and python language both are present
					print(e.options)
					v.Speak("Too many results for this keyword.Please be more specific and try again")
					continue
				except wikipedia.exceptions.PageError as e: #page not found
					print('The page does not exist')
					v.speak('The page does not exist')
					continue
			elif sc in message: #speak deep search any query
				try:
					words=message.split()
					del words[0:2]
					st=' '.join(words)
					scq=c1.query(st) #takes any query input 
					sca=next(scq.results).text #next iterates over  many results come as output and converts them in text format
					print('The answer is: '+str(sca))
					#url='http://www.wolframalpha.com/input/?i='+ st #used when u don't have access key of ur wolframa I'd it directly take access from wolframalpha site
					#webbrowser.open(url)
					v.speak('The answer is: '+str(sca))
				except StopIteration:
					print('Your question is ambiguous.Please try again')
					v.speak('Your question is mbiguous.Please try again')
				except Exception as e:
					print(e)
					v.speak(e)
			elif paint in message.lower():
				os.system('mspaint')
			elif rdds in message: #read the text which is currently present in clipboard or whenever we copied something is automatically saved to clipboard and it will apeak out that text
				print('Reading ur text')
				v.Speak(pyperclip.paste()) # to paste text
			elif sav in message:
				with open('path to your text file','a') as f: #for saving text in clipboard in a file pathname should be given as a
					f.write(pyperclip.paste())
				print("Saving your text to file")
				v.speak("saving your text to file")
			elif bkmk in message:
				shell.SendKeys("^d")
				v.speak("Page bookmarked")
			elif keywd in message:
				print('')
				print('Say " '+google+' "to return a google search')
				print('Say " '+acad+' "to return a google scholar search')
				print('Say " '+sc+' "to return a Wolfram Alpha query')
				print('Say " '+wkp+' "to return a Wikipedia page')
				#print('Say " '+book+' "to return an Amazon book search')
				print('Say " '+rdds+' "to read the text you have highlighted and ctrl+C (coped to clipboard)')
				print('Say " '+sav+' "to save the text you have highlighted and Ctrl+c-ed(copied to clipboard )to a file')
				print('Say " '+bkmk+' "to bookmark the page you are currently reading in your browser')
				print('Say " '+vid+' "to return video results for your query')
				print('For more general questions,ask them naturally and I wil do my best to find a good answer')
				print('Say '+stoplst+' to shut down')
				print('')
			elif vid in message: #speak 'video for python'
				words=message.split()
				del words[0:2]
				st=' '.join(words)
				print('Video Results for: '+str(st))
				url='https://www.youtube.com/results?search_query='+st
				webbrowser.open(url)
				v.Speak('Video Results for: '+str(st))
			elif wtis in message: #search for things that r present in wolframa like 5+10 
				try:
					scq=c1.query(message) #c1 object of wolframa, it takes query from user and finds the result from there
					sca=next(scq.results).text
					print('The answer is: '+str(sca))
					#url='http://www.wolframalpha.com/input/?i='+st
					#webbrowser.open(url)
					v.Speak('The answer is: '+str(sca))
				except UnicodeEncodeError:
					v.Speak('The answer is: '+str(sca))
				except StopIteration: #search for things which r not in wolframa like python languauge
					words = message.split()
					del words[0:2]
					st=' '.join(words)
					print('Google Results for: '+str(st))
					url='http://google.com/search?q='+st #directly search the url for entered message
					webbrowser.open(url)
					v.Speak('Google Results for:')
			elif wtar in message: #search for things that r present in wolframa like 5+10 
				try:
					scq=c1.query(message) #c1 object of wolframa, it takes query from user and finds the result from there
					sca=next(scq.results).text
					print('The answer is: '+str(sca))
					#url='http://www.wolframalpha.com/input/?i='+st
					#webbrowser.open(url)
					v.Speak('The answer is: '+str(sca))
				except UnicodeEncodeError:
					v.Speak('The answer is: '+str(sca))
				except StopIteration: #search for things which r not in wolframa like python languauge
					words = message.split()
					del words[0:2]
					st=' '.join(words)
					print('Google Results for: '+str(st))
					url='http://google.com/search?q='+st #directly search the url for entered message
					webbrowser.open(url)
					v.Speak('Google Results for:')
			elif whis in message: #search for things that r present in wolframa like 5+10 
				try:
					scq=c1.query(message) #c1 object of wolframa, it takes query from user and finds the result from there
					sca=next(scq.results).text
					print('The answer is: '+str(sca))
					#url='http://www.wolframalpha.com/input/?i='+st
					#webbrowser.open(url)
					v.Speak('The answer is: '+str(sca))  #search for things which r not in wolframa like python languauge
				
				except StopIteration:
					try:
						words = message.split()
						del words[0:2]
						st=' '.join(words)
						wkpres=wikipedia.summary(st, sentences+2)
						print('\n'+ str(wkpres) +'\n')
						v.Speak(wkpres)
					except UnicodeEncodeError:
						v.Speak(wkpres)
					except:
						words = message.split()
						del words[0:2]
						st=' '.join(words)
						print('Google Results(last exception) for: '+str(st))
						url='http://google.com/search?q='+st #directly search the url for entered message
						webbrowser.open(url)
						v.Speak('Google Results for:'+str(st))

		except Exception as e:
			print(e)
			break
		finally:
			break