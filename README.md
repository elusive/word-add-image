# word-add-image
Command line utility to add image to open word document.

## Usage
Should be called with two arguments:
- Path to image
- Alt text for image

```
word-add-diagram.exe "c:/images/image.jpg" "Diagram definition as alt text."
```

This is very much a first pass but functional. The following are the possible
return values and each is returned for the described situation or exception.
```
Success = 0;
EmptyImagePath = 1;
ImagePathNotFound = 2;
EmptyAltText = 3;
NoOpenWordDocument = 4;
FailedToAddImage = 5;
```
