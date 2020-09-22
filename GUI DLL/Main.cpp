/***********************************************************************
*This application(ProcessXP) and its components(ProcXPGUI.dll, sndServer.dll) were explicitly 
*developed for PSC(Planet Source Code) Users as Open Source Projects.
*This code and the code of its components are property of their author.
*
*If you compile this standard dynamic link library you may not redistribute it.
*However, you may use any of this code in you're own dynamic link library or application(s).
*
*Alex Smoljanovic, Salex Software (c) 2001-2003
*salex_software@shaw.ca
***********************************************************************/


#include <windows.h>
#include <stdio.h>
//#include <string.h>
#include <Wingdi.h>
//header files to include

byte r,g,b,oR,oG,oB;
//declare variables r, g, b, oR, oG, oB as byte data structures

const long Black2White = 0;
const long White2Black = 1;
const long NOFADE = 0;
const long EQUALITYFADE = 1;
const long UNEQUALITYFADE = 2;
//constant flags used while drawing gradients...

void DoEvents(void);
void UpdateColor(int, int, bool);
int _stdcall VerticalGrad(HDC, int, int, long, bool);
void FadePixels(int, int);
int _stdcall HorizontalGrad(HDC, int, int, long, long, long);
void gUpdateColor(int, int);
void gFadePixels(int, int);
int _stdcall HRadialGrad(HDC, int, int, long);
//Forward function declarations...



void DoEvents(void) 
{ 
 MSG Msg; 
  while(PeekMessage(&Msg,NULL,0,0,PM_REMOVE))
  {
   if (Msg.message == WM_QUIT)break;
    TranslateMessage(&Msg); 
     DispatchMessage(&Msg);
  } 
}
//This function(DoEvents) from Jared Brundi, thanks Jared.



void UpdateColor(int cPos, int MaxPos, bool b2w)
{
/*
This function will update the global color bytes based upon the drawing position.
*/
 if (b2w==Black2White){
  //Flag Black2While specifies the direction in which to draw the gradient.
  //If b2w evaluates to Black2While the color specified will be drawn first, fading to white as
  //opposed to White2Black, where white is the first color drawn, fading to the system color specified
  r = oR + (double(double(cPos) / double(MaxPos)) * (255 - oR));
  //initialize variable r with oR(Initial Red Color) with the equivelant of the percent drawn to the percent of oR(initial red)'s value out of 255
  //so if 50% is drawn, then r will now evaluate 50% more intense(white) than the inital red color
   g = oG + (double(double(cPos) / double(MaxPos)) * (255 - oG));
   //..
	b = oB + (double(double(cPos) / double(MaxPos)) * (255 - oB));
	//..
 }
 else{
  r = 255 - (double(double(cPos) / double(MaxPos)) * (255 - oR));
  //the same, except if 50% is drawn, then r will now evaluate to 50% of the initial red color
   g = 255 - (double(double(cPos) / double(MaxPos)) * (255 - oG));
   //..
	b = 255 - (double(double(cPos) / double(MaxPos)) * (255 - oB));
	//..
 }

}

int _stdcall VerticalGrad(HDC hDc, int Height, int Width, long SysColInd, bool b2w)
{
 //this function is exported as an STD Call; see function Definition file (.def)
 try{
 //try catch statement, to catch any error that may occur, allthough no error will most likely occur...
 DWORD tmpColor = GetSysColor(SysColInd); //return the System defined color specified by the SysColInd argument...
  oR = GetRValue(tmpColor); //call GetRValue macro to return the Red color value of the tmpColor variable
   oG = GetGValue(tmpColor); //..
    oB = GetBValue(tmpColor); //..
     int cx,cy; //declare variables cx and cy as integers
      for(cy=0; cy <= Height; cy++){
	  //loop; initialize cy to 0, 
	  //loop while cy is less than or equal to the height of the device context handle to which we are drawing
	  //increment cy
       DoEvents(); //Yield execution so that other procedures processing asynchronously can process
        UpdateColor(cy, Height, b2w); //see UpdateColor for more info..
	     for(cx=0; cx <= Width; cx++){
		 //loop; initialize cx to 0
		 //loop while cx is less than or equal to the width of the hdc(device context handle) to which we are drawing the gradient
	      SetPixel(hDc, cx,cy, RGB(r,g,b));
		  //set the specified pixel of the hdc to the color reference returned by the RGB macro
		 }
	  }
       return 1; //function was successful, return 1
 }
 catch(void){
  return 0; //return 0; an error occured
 }
}


void FadePixels(int cPos, int MaxPos)
{
  r = r + (double(double(cPos) / double(MaxPos)) * (255 - r));

   g = g + (double(double(cPos) / double(MaxPos)) * (255 - g));
   //..
    b = b + (double(double(cPos) / double(MaxPos)) * (255 - b));
	//..
}


int _stdcall HorizontalGrad(HDC hDc, int Height, int Width, long SysColInd, long b2w, long fade)
{
 try{
 DWORD tmpColor = GetSysColor(SysColInd); //return the System defined color specified by the SysColInd argument...
  oR = GetRValue(tmpColor); //initialize oR with the red value of tmpColor
   oG = GetGValue(tmpColor); //initialize oG with the green value of tmpColor
    oB = GetBValue(tmpColor); //initialize oB with the blue value of tmpColor
     int cx,cy;
      for(cx=0; cx <= Width; cx++){
	  //loop;
	  //initialize cx to 0
	  //loop while cx is less than or equal to the width specified in the width argument
	  //increment cx,,,
        DoEvents(); //Yield execution to asynchronously processing procedures
    	 UpdateColor(cx, Width, b2w); //See UpdateColor for more info
		  if (fade == EQUALITYFADE){
			FadePixels(cx, Width); //if fade evaluates to Equality; see FadePixels function for more info...
		  }
	      for(cy=0; cy <= Height; cy++){
		   if (fade == UNEQUALITYFADE){
			//if fade evaluates to UnEqualityFade then...
			FadePixels(cy, Height); //See FadePixels function for more info...
		   }
	        SetPixel(hDc, cx,cy, RGB(r,g,b)); //RGB macro returns a color reference
			//set the specified(x,y) pixel's color of the specified device context handle
		  }
	  }
       return 1; //return 1, function was successful
 }
 catch(void){
  return 0; //return 0, an error occured
 }
}



void gUpdateColor(int cPos, int MaxPos)
{
 //function updates the r, g, and b values based upon the percent drawn
  r = 255 - (double(double(cPos) / double(MaxPos)) * (255 - oR));
  //initialize r to the value of (White - %Drawn) multiplied by (White - Initial Red)
   g = 255 - (double(double(cPos) / double(MaxPos)) * (255 - oG));
   //...
	b = 255 - (double(double(cPos) / double(MaxPos)) * (255 - oB));
	//...
}


void gFadePixels(int cPos, int MaxPos)
{
//This function will increase the intesity(whiteness) of the pixels on yet an additional angle;
//a diagonal angle(the angle is equivelant from the top left corner of the dimensions being drawn to the lower right relatively) progressively

  r = 255 - (double(double(cPos) / double(MaxPos)) * (255 - r));
  //r is incremented by (Percent Drawn) multiplied by (white - current red)
  /*
	so while drawing a vertical gradient with b2w evaluating to Black2White(System Color to White;Top to bottom)
	the following will occur; below 10 equals the full intensity of the System Color, 0 equals white

    ___________
   |10 10 9	9 | Notice that this remains a vertical gradient
   |8   8 7 7 | yet the pixels intesity diagonally are also increased.
   |6   6 5 5 | This is similar to a diagonal gradient, but not a true diagonal gradient
   |4	4 3 3 | for a diagonal gradient is identical to a horizontal or vertical gradient
   |2   2 1 1 | where the color gradient occurs on only one angle
   |0   0 0 0 |
   ------------
  */
   g = 255 - (double(double(cPos) / double(MaxPos)) * (255 - g));
    b = 255 - (double(double(cPos) / double(MaxPos)) * (255 - b));
}

int _stdcall HRadialGrad(HDC hDc, int Height, int Width, long SysColInd)
{
 try{
 DWORD tmpColor = GetSysColor(SysColInd);//Return the system defined color
  oR = GetRValue(tmpColor); //return red color of tmpColor
   oG = GetGValue(tmpColor); //return green color of tmpColor
    oB = GetBValue(tmpColor); //return blue color of tmpColor
     int cx,cy;
      for(cx=Width; cx >= 0; cx--){
        DoEvents();
    	 gUpdateColor(cx, Width); //See gUpdateColor function...
	      for(cy=Height; cy >=0 ; cy--){
		   gFadePixels(cy, Height); //See gFadePixels function...
	        SetPixel(hDc, cx,cy, RGB(r,g,b));
			//set the specified pixel's color
		  }
	  }
       return 0;
 }
 catch(int err){
  return err;
 }
}
