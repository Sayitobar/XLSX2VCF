# XLSX2VCF
Converts XLSX documents into VCF (Variant Call Format) - remastered for JDK1.8  
The code is actually [from Apache](https://svn.apache.org/repos/asf/poi/trunk/poi-examples/src/main/java/org/apache/poi/examples/xssf/eventusermodel/XLSX2CSV.java), but it didn't work as intended and was old. I rewrote some parts and wrote a function to make things easy!
- Supports large files (thanks to XSSF)
- Very fast conversion
- Has "temporary file" feature
- Remastered & easy to use!

## Usage
Just read `Example.java` file. There is a really easy to use converter function as example. 
![ss](https://user-images.githubusercontent.com/95364352/233811494-2cd7b5d6-1f86-4075-93cb-5a099cfff1a4.png)
- If you want to convert XLSX file to VCF file temporarily (for fast processing etc.), there is a custom parameter `temporary` that deletes converted files after use.

## Dependencies
- You need to download some third party JARs to work with XLSX files, Apache POI to be more precise.
- [This is the link](https://poi.apache.org/download.html#POI-5.2.3) to the download website, but it's a bit complicated. So I'll upload JAR files I use just to make stuff more easy.
