# Convert text file to powerpoint.
This program is intended to put texts in the .txt file on powerpoint slides. 

## Fuctions
- Create slide format automatically.
- indents first sentence of paragraph.
- line breaks without splitting words.
- divides texts into each slide according to designated number of lines.
- Each line is splitted depending on the given length.
- Once you assigned
## How to use
1. There are several variables to assign.
* font size
* font name
* line width
* The number of lines per slide
* Slide width and height
* Slide/text box margin on left, right, top and bottom
* line spacing
* Text file directory
* Powerpoint file directory(The output pptx file will be stored in this path)
2. Once you finish filling all inputs, run the program!
3. The output is going to be like below.

### _Original text file_
>텍스트 파일을 ppt에 예쁘게 얹는 프로그램을 만들었다. 페이스북에 올리는 글을 인스타그램에도 올린다. 글은 파워포인트를 거쳐 이미지로 변환한다. 각 문단의 첫 줄은 들여 쓰기 하고 13줄씩 쪼개서 슬라이드 별로 앉힌다. 똑같은 작업을 글 쓸 때마다 반복하는 게 귀찮고 간단한 작업도 사람 손으로 하다 보니 실수할 때가 있었다. 
>python의 pptx 라이브러리를 끌어와 만들었는데 많이 쓰이는 모듈이 아니어서 그런지 ppt의 모든 기능을 지원하지 않는다. 줄바꿈 할 때 글자 단위가 아닌 단어 단위로 끊어서 넘겨 주는 기능이 없다.(단어 단위로 줄넘김 하는 게 더 가독성이 좋다. 나는 그렇다.) 그래서 글을 단어별로 다 쪼개고, 쪼갠 단어의 길이의 합이 일정 길이가 되면 한 줄로 만들었다. 슬라이드 당 13줄씩 앉히는 것까지 성공! ppt를 jpg로 저장하는 코드는 인터넷에 돌아다니는 걸 따다 썼다. txt 파일을 저장하고 실행하면 지정한 폴더에 슬라이드 몇 장이 다소곳이 저장되어 있다. 흐흐흐
>프로그래밍은 지겨운 일을 자동으로 할 수 있게 도와준다. 내가 필요한 걸 만드니 몰랐던 기능도 배울 의지가 생긴다. 흥미와 새로운 배울 거리가 적절히 섞인 다음 아이템을 찾아보자! 
### _Created Powerpoint slide_
![Slide1](https://user-images.githubusercontent.com/84579416/128120738-2ea9a5fc-0c86-425b-a613-3629792a8999.jpg)
![Slide2](https://user-images.githubusercontent.com/84579416/128120746-1e732d31-ed91-4613-a5c7-ad933b20cc9a.jpg)
