function write3mf(filename , vertices , faces, colors)
% WRITE3MF: Writes 3mf 3D file from vertices, faces and optionally
% colors.
% 3mf file is the "3D Manufacturing Format" of file that will allow design 
% applications to send full-fidelity 3D models to a mix of other 
% applications, platforms, services and printers. Full description and
% specifications can be found here: https://3mf.io/
% I wrote this function because 3mf files can be imported in Powerpoint 365
% ProPlus 2016 as 3D objects which can be manipulated live.
% 
% This function creates a simple 3mf file structure, embedding the
% vertices, faces and colors provided.
%
% Syntax: write3mf(filename , vertices , faces, colors)
%
%   Inputs:
%       filename - output 3mf file
%       vertices - list of unique vertices, Nx3 coordinates
%       faces    - connection matrix, Mx3 triangular mesh
%       colors   - (optional) RGB color for each vertex, Nx3 values (between 0 and 1 OR
%                  between 0 and 255)
%
% Demo mode: write3mf(filename , 'demo')
% will create a demo 3mf file
%
% Example:
% 	vertices = [0 0 0; 10 0 10; 10 10 0; 0 10 10];
%	faces    = [1 3 4; 2 3 4; 1 2 4; 1 4 3;];
%   write3mf('D:\temp\pyramid_vertexcolor.3mf' , vertices , faces)

    
    if nargin < 2
        error('Not enough input arguments. Usage: write3mf(filename , vertices , faces, colors)');
    elseif nargin < 3 && ischar(vertices) && strcmpi(vertices , 'demo')
        % Demo mode
        [vertices , faces, colors] = prepareDemo();
    elseif nargin < 4
        % No colors provided
        colors = [];
    end
    
    % Check if input is valid
    checkInput(filename , vertices , faces, colors);
    
    % Temporary files
    files = tempFiles(filename); 
    writeTempFiles(files);    
    
    % 3D model
    write3DModel(files , vertices , faces, colors);

    % Builds 3mf file
    package3mf(files)

    % Remove temp files
    cleanTempFiles(files)        
    


function hex = rgb2hex(rgb)
% rgb2hex function is adapted from Chad Greene's function, which can be
% found on MathWork's File Exchange at this address:
% https://fr.mathworks.com/matlabcentral/fileexchange/46289-rgb2hex-and-hex2rgb

    if max(rgb(:))<=1
        rgb = round(rgb*255); 
    else
        rgb = round(rgb); 
    end
    hex(:,1:6) = reshape(sprintf('%02X',rgb.'),6,[]).'; 
    
    
    
function checkInput(filename , vertices , faces, colors)    

    if size(vertices,2) ~= 3
        error('Error: Vertices should be a Nx3 matrix')
    elseif size(faces,2) ~= 3
        error('Error: Mesh should be triangular; faces should be a Mx3 matrix')
    elseif max(faces(:)) > size(vertices,1)
        error(['Error: Some faces indicate a non-existing node number ' num2str(max(faces(:)))]);
    elseif ~isempty(colors) && size(colors,2) ~= 3
        error('Error: colors variable should be empty or a Nx3 matrix of RGB values for each vertex')
    elseif ~ischar(filename)
        error('Error: filename should be a string pointing to the output file')
    end


    
function files = tempFiles(filename)
% write3mf must write some temporary files in a specific path structure,
% reproducing the structure of the 3mf file. 

    % Try using the OS temporary path; use try/catch because sometimes
    % admin privileges are needed to write in the temporary folder  
    try
        files.temppath = tempdir();
        files.temppath = [files.temppath , 'write3mf' , filesep];
        tmp_file = [files.temppath 'test.txt'];
        fid = fopen(tmp_file,'w');
        fprintf(fid , '%s' , 'Test writing');
        fclose(fid);
        delete(tmp_file);
        
    catch
        % If temporary folder does not work, use local path
        
        files.temppath = [mfilename('fullpath') , '_temp', filesep];
        if exist(files.temppath , 'dir')
            error(['Error: write3mf must write some temporary files but the ' ...
                   'temporary folder already exists. Please start Matlab with '...
                   'admin privileges so that the OS temporary folder can be used, ' ...
                   'or delete this folder before proceeding: '...
                   files.temppath]);
        end
        
    end

    % Structure and files for the ZIP file that will be packaged withing
    % the 3mf output file
    files.content_types = [files.temppath , '[Content_Types].xml'];
    files.p_rels = [files.temppath , '_rels', filesep];
    files.p_3D   = [files.temppath , '3D', filesep];
    files.rels   = [files.p_rels '.rels'];
    files.model  = [files.p_3D , '3dmodel.model'];
    files.zip = [files.temppath , '3mf.zip'];
    files.output = filename;
    
    mkdir(files.temppath)
    mkdir(files.p_rels)
    mkdir(files.p_3D)

    
    
    
function package3mf(files)    
% Create zip file and renames it to the output 3mf file
    
    zip(files.zip , {'[Content_Types].xml' , ['_rels', filesep, '.rels'] , ['3D' filesep '3dmodel.model']},files.temppath);
    movefile(files.zip , files.output , 'f')


    
function cleanTempFiles(files)        
% Removes temp files and folders

    delete(files.content_types , files.rels, files.model)
    rmdir(files.p_rels)
    rmdir(files.p_3D)
    rmdir(files.temppath)


    
function writeTempFiles(files)    

    % .rels file
    fid = fopen(files.rels , 'w');
    fprintf(fid , ['<?xml version="1.0" encoding="UTF-8"?>'...
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'...
        '<Relationship Target="/3D/3dmodel.model" Id="rel0" Type="http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel" />'...
        '</Relationships>\n']);
    fclose(fid);

    
    fid = fopen(files.content_types , 'w');
    fprintf(fid , ['<?xml version="1.0" encoding="UTF-8"?>'...
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'...
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />'...
        '<Default Extension="model" ContentType="application/vnd.ms-package.3dmanufacturing-3dmodel+xml" />'...
        '</Types>\n']);
    fclose(fid);
 
    
    
    
function write3DModel(files , vertices , faces, colors)
% Writes the 3D model proper

    fid = fopen(files.model , 'w');
    
    % Write header
    fprintf(fid , '<?xml version="1.0" encoding="UTF-8"?>\n');
    fprintf(fid , '<model unit="millimeter" xml:lang="en-US" xmlns:m="http://schemas.microsoft.com/3dmanufacturing/material/2015/02" xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02">\n');
    fprintf(fid , '\t<metadata name="CreationDate">2015-08-12</metadata>\n');
    fprintf(fid , '\t<metadata name="Description">Created by write3mf for Matlab V00</metadata>\n');
    fprintf(fid , '\t<resources>\n');
    
    
    % Prepare and write colors for each vertex, if present
    if ~isempty(colors)
        [unique_colors , ~, idx_colors] = unique(colors , 'rows');
        unique_colors = rgb2hex(unique_colors);

        fprintf(fid , '\t\t<m:colorgroup id="2">\n');
        for k = 1 : size(unique_colors , 1)
            fprintf(fid , '\t\t\t<m:color color="#%s" />\n' , unique_colors(k,:));        
        end
        fprintf(fid , '\t\t</m:colorgroup>\n');
    end

    
    % Start object
    fprintf(fid , '\t\t<object id="1" name="write3mf_3Dmodel" type="model">\n');
    fprintf(fid , '\t\t\t<mesh>\n');
    
    % Vertices
    fprintf(fid , '\t\t\t\t<vertices>\n');
    for k = 1 : size(vertices , 1)
        fprintf(fid , '\t\t\t\t\t<vertex x="%.2f" y="%.2f" z="%.2f" />\n' , vertices(k,:));        
    end    
    fprintf(fid , '\t\t\t\t</vertices>\n');
    
    % Triangles
    fprintf(fid , '\t\t\t\t<triangles>\n');
    if isempty(colors)
        for k = 1 : size(faces , 1)
            fprintf(fid , '\t\t\t\t\t<triangle v1="%u" v2="%u" v3="%u" />\n' , faces(k,:) - 1);        
        end        
    else
        for k = 1 : size(faces , 1)
            row_colors = idx_colors(faces(k,:))' - 1;
            row_faces = faces(k,:) - 1;
            fprintf(fid , '\t\t\t\t\t<triangle v1="%u" v2="%u" v3="%u" pid="2" p1="%u" p2="%u" p3="%u" />\n' , row_faces , row_colors);        
        end        
    end
    fprintf(fid , '\t\t\t\t</triangles>\n');
    
    
    % Footer
    fprintf(fid , '\t\t\t</mesh>\n');
    fprintf(fid , '\t\t</object>\n');
    fprintf(fid , '\t</resources>\n');
    fprintf(fid , '\t<build>\n');
    fprintf(fid , '\t\t<item objectid="1" />\n');
    fprintf(fid , '\t</build>\n');
    fprintf(fid , '</model>\n');
    
    fclose(fid);    
    
    
    
function [vertices , faces, colors] = prepareDemo()
% This demo reproduces the "pyramid_vertexcolor.3mf" example that can be found at:
% https://github.com/3MFConsortium/3mf-samples/blob/master/materials/pyramid_vertexcolor.3mf
    
    vertices = [0 0 0; 10 0 10; 10 10 0; 0 10 10];
    faces = [0 2 3; 1 2 3; 0 1 3; 0 3 2;] + 1;
    colors = [255 0 0; 0 0 255; 0 255 0; 255 255 255];
    idx_colors = [1 3 2 4];
    colors = colors(idx_colors,:);

    figure
    patch('Faces', faces, 'Vertices',vertices,'FaceVertexCData', colors,...
          'EdgeColor','None', 'FaceColor', 'interp', 'CDataMapping' , 'scaled')
    title('Creating 3mf file of this object...')
    axis equal;  grid on;
    set(gca, 'XTickLabel' , '', 'YTickLabel' , '', 'ZTickLabel' , '')
    campos([15 , -70 , 45]);
    
    
    