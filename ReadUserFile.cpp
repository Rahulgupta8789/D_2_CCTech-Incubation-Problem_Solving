/*
1.Read a user defined file. 
count the following:
  i>    numbers of characters in the file.
  ii>   number of white spaces.
  iii>  number of words in the file.
  iv>   number of words starting with each vowel in the file.
  v>    number of punctuation marks in the file.
  vi>   number of occurences of a user defined string in the text in the file.

*/

#include <iostream>
#include <vector>
#include <string>
#include <sstream>
#include <fstream>
#include <map> 

using namespace std;

void fileProcessing(string path) 
{
    vector<string> words;                // store total words
    map<char, int> vowel_words;          // store each word start with vowel
    int whitespace=0; 
    int punctuation_marks=0; 
    string line; 
    ifstream fin;                         // declare input file stream object
    string word; 
    int total_character = 0; 

    fin.open(path);                       // open file 
    if(fin.is_open())                     // check file is available or not
    {                   

        // read file line by line
        while(getline(fin, line)) 
        {
            for(int i = 0; i < line.length(); ++i) 
            {

                // count white space
                if(line[i] == ' ' || line[i] == '\t')
                {
                    whitespace++;
                } 

                // count punctuation marks
                if(line[i] == '!' || line[i] == ',' || line[i] == ';' || line[i] == '.' || line[i] == '?' 
                  || line[i] == '-' || line[i] == '\'' || line[i] == '\"' || line[i] == ':' || line[i] == '/') 
                  {
                    punctuation_marks++;
                  }
            }

            // store each word from each line into vector
            istringstream ss(line);
            while( ss >> word) 
            {
                words.push_back(word);
            }
        }

        for(auto s : words) 
        {
            
            total_character = total_character + s.length();    // calculate total number of character

            // count number or word starts with each vowel
            char temp = toupper(s[0]);
            if(temp == 'A' || temp == 'E' || temp == 'I' || temp == 'O' || temp == 'U') 
            {
                if(vowel_words.find(temp) == vowel_words.end()) 
                {
                    vowel_words.insert(make_pair(temp, 1)) ;
                }
                else 
                {
                    vowel_words[temp]++;
                }
            }
        }

        // display result
        cout << endl << "Total number of character :  " << total_character;
        cout << endl << "Total number of white space : " << whitespace;
        cout << endl << "Total number of words : " << words.size();
        
        // display number of words start with each vowel
        cout << endl << "Number of words start with each vowel : ";
        for(auto &it : vowel_words) 
        {
            cout  << endl<< it.first << " --> " << it.second;
        }

        cout << endl << "Total number of punctuation marks : " << punctuation_marks;

        // find input string is present or not
        string input;
        cout << endl << "Enter any string to search in text file : ";
        getline(cin, input);
        int count = 0;
        for(auto s : words) 
        {
            if(s == input) 
            {
               count++;
            }
        }
        cout << endl << "Total number of occurence of string \"" << input << "\" : " << count;   
    }
    else 
    {
        cout << "Specified file not found !";
    }
}

int main() 
{
    string path;
    cout << endl << "Enter the Name of File : " ;
    getline(cin, path);
    fileProcessing(path);
    return EXIT_SUCCESS;
} 