# XLSX2VCF
Converts XLSX documents into VCF (Variant Call Format) - remastered for VCF with JDK1.8  
The code is actually [from Apache](https://svn.apache.org/repos/asf/poi/trunk/poi-examples/src/main/java/org/apache/poi/examples/xssf/eventusermodel/XLSX2CSV.java) for converting XLSX to CSV. However, it didn't work as intended. I rewrote some parts from scratch, adapted it for Variant Call Format and wrote a function to make things easy!
- Supports large files (thanks to XSSF)
- Very fast conversion
- Has "temporary file" feature
- Remastered for VCF & easy to use!

## Usage
Just read `Example.java` file. There is a really easy to use converter function as example. 
![ss](https://user-images.githubusercontent.com/95364352/233811494-2cd7b5d6-1f86-4075-93cb-5a099cfff1a4.png)
- If you want to convert XLSX file to VCF file temporarily (for fast processing etc.), there is a custom parameter `temporary` that deletes converted files after use.

## Dependencies
- You need to download some third party JARs from Apache POI to work with XLSX files.
- [This is the link](https://poi.apache.org/download.html#POI-5.2.3) to the download website, but may be a bit complicated. So I'll upload JAR files I use just to make stuff more easy.
