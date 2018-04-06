# write3mf  
WRITE3MF: Writes 3mf file from vertices, faces and optionally colors.  
3mf file is the "3D Manufacturing Format" of file that will allow design applications to send full-fidelity 3D models to a mix of other applications, platforms, services and printers. Full description and specifications can be found here: https://3mf.io/  
I wrote this function because 3mf files can be imported in Powerpoint 365 ProPlus 2016 as 3D objects which can be manipulated live!
  
This function creates a simple 3mf file structure, embedding the vertices, faces and colors provided.  

Syntax: `write3mf(filename , vertices , faces, colors)` 
 
### Inputs:  
>       filename - output 3mf file  
>       vertices - list of unique vertices, Nx3 coordinates  
>       faces    - connection matrix, Mx3 triangular mesh  
>       colors   - (optional) RGB color for each vertex, Nx3 values (between 0 and 1 OR  
                 between 0 and 255)  
 
### Demo mode: 
`write3mf(filename , 'demo')` will create a demo 3mf file  
 
### Example:  
```
vertices = [0 0 0; 10 0 10; 10 10 0; 0 10 10];  
faces    = [1 3 4; 2 3 4; 1 2 4; 1 4 3;];  
write3mf('D:\temp\pyramid_vertexcolor.3mf' , vertices , faces)
```

