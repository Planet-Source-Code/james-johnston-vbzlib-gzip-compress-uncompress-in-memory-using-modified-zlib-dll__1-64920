<div align="center">

## VBZLib \- GZip Compress/Uncompress In Memory using Modified ZLib DLL


</div>

### Description

This is a class wrapper and a modified version of the ZLib DLL (http://www.zlib.net). ZLib is used for compressing/uncompressing data using the famous Deflate algorithm.

It will also read/write the zlib and gzip headers from/to the data when compressing/uncompressing. I modified the standard zlib compress/uncompress function calls to allow you to pass a parameter to specify the header type of the data. Either zlib, gzip, auto or deflate with no header.

You can download "vbzlib1.dll" from the following URL.

http://www.techknowpro.com/vb/vbzlib/vbzlib1.dll

I needed this because I couldn't find any other sample code to uncompress a string containing gzip data. Specifically when downloading gzip Content-Encoded data from a web server.

With this you can both compress and uncompress string data easily specifying what ever header you like.

I hope this helps others. I did alot of searching and didn't find anything like it for VB.

If you like or use this please vote for me. :)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2006-04-03 23:00:02
**By**             |[James Johnston](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/james-johnston.md)
**Level**          |Advanced
**User Rating**    |4.7 (71 globes from 15 users)
**Compatibility**  |VB 6\.0
**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__1-49.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[VBZLib\_\-\_G198518442006\.zip](https://github.com/Planet-Source-Code/james-johnston-vbzlib-gzip-compress-uncompress-in-memory-using-modified-zlib-dll__1-64920/archive/master.zip)








