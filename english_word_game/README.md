English Word Game

I created this project, English Word Game, with the intent to use Excel formulas to build a fun and practical tool that converts English words into numerical scores. By doing so, I reinforced key Excel skills like text manipulation, lookups, logical testing, and cell referencing—all through a creative challenge.

To start, I built a lookup table listing each letter of the alphabet (A through Z) alongside its corresponding numeric value (1 through 26). This reference table sits in cells A4:B29 and serves as the foundation for the scoring logic. It allows me to match any letter from a given word to its assigned numeric value using lookup functions later in the process.

Next, I set up the main input cell in C5, where I can type any English word I want to score. For example, entering the word EXCEL triggers a series of formulas that break the word into individual letters, find each letter’s numeric value, and calculate the total score.

To display the alphabet dynamically, I used the formula =CHAR(64+B4). This takes advantage of Excel’s character encoding system, where 65 corresponds to “A.” By offsetting the number by 64, I can easily generate letters A–Z as I fill the formula down through the rows. This gave my worksheet a dynamic alphabet reference without typing each letter manually.

To extract letters from the input word, I used the MID() function:

=MID($C$5, B4, 1)


This formula isolates a single character from the word in C5. The $C$5 reference is absolute so that it stays fixed, while B4 changes with each row to extract the next letter. I love how this technique makes the extraction process both clean and scalable—it works for any word length up to 26 characters.

After pulling each letter, I used a lookup formula to assign its numeric value:

=IFERROR(LOOKUP(F4, $A$4:$B$29), "")


This line looks up the extracted letter (from column F) in the alphabet table and returns the corresponding number. To make the worksheet more user-friendly, I wrapped the formula in an IFERROR() function so that empty cells don’t show distracting error messages.

Once each letter’s numeric score was calculated, I summed them all using:

=IF(SUM(G:G)=0, "", SUM(G:G))


This ensures that if no word is entered (and therefore no valid values exist), the cell remains blank. Otherwise, it calculates the total “word score” by summing all the numeric letter values.

While building this game, I practiced several important Excel principles. I worked intentionally with relative and absolute references—keeping $C$5 fixed while allowing other cells to update dynamically. I also gained deeper understanding of text functions like MID(), explored character encoding through CHAR(), and handled potential lookup errors with IFERROR(). Finally, I reinforced aggregation logic by using SUM() to compute total scores.

Here’s an example result:
If I type EXCEL in cell C5, the extracted letters (E, X, C, E, L) convert to values (5, 24, 3, 5, 12). Summing these gives a final score of 49. This simple output represents the project’s purpose—turning text into meaningful numeric data using layered Excel logic.

To make the project more engaging, I also experimented with a few optional enhancements. I applied conditional formatting to highlight high-value letters, added data validation to restrict input to alphabetic characters, and even explored creating a Scrabble-style scoring system. These tweaks made the spreadsheet more interactive while helping me apply Excel’s visual and validation tools.

This project draws on three main files:

Mastering_Excel_Through_Projects.pdf — the instructional material (pages 25–31).

formula.pdf — my quick reference sheet listing all formulas.

english_word_game.xlsx — the actual Excel file where I implemented and tested everything.

Completing this project taught me how to combine multiple Excel functions into a cohesive system. It’s a perfect example of how seemingly simple functions—like MID, CHAR, LOOKUP, and IFERROR—can work together to build an elegant, logic-driven solution. More importantly, it showed me how Excel can be used creatively, not just analytically.

In short, building the English Word Game helped me strengthen my understanding of text manipulation, lookups, and conditional logic in Excel, while also reinforcing good formula design and debugging habits. It’s a small but powerful demonstration of how structured thinking and attention to detail can turn a spreadsheet into a functional, interactive model.