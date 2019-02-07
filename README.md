# xlreport
Generate report from a xlsx excel template. Just set some mark in template, xlreport will parse the document, and fill the cell with a dictionary. Supporting `Numpy Array` as image, and `Pandas Dataframe` as table.



## Personal Information Demo

![](http://idoc.imagepy.org/demoplugin/33.png)

<div align=center>the template document</div><br>

```python
    wb = load('Personal Information.xlsx')
    info, keys = parse(wb)
    
    # prepare the data
    img = np.random.randint(0, 255, (100,100), dtype=np.uint8)
    data = {'Name':'YX Dragon',
            'Photo':img,
            'Sex':'Male',
            'Age':'30',
            'Height':170,
            'Weight':75.0,
            'Like_Sport':True}
    
    fill_value(wb, info, data)
    repair(wb) # to repair some bug of openpyxl
    
    wb.save(osp.join(direc, 'person.xlsx'))
```

then we can got a filled document, here we save it as a pdf document.

![](http://idoc.imagepy.org/demoplugin/34.png)

## Rule of Template

*Here is the rule of make a template, please set your `page size`, suchas A4, A5. Then change the view to `page view`, then the cell's size could be set in cm.*

**General Rule**

`{type Var_Name = Default Value # note}`，any mark should be in `{}`, and must have `type` and `var_name`, some var must have `default value`.

**Basic Parameter**

`int,float,str,txt,bool`：these vars has the same format, `str` for a demo，`{str Name = YX Dragon # please input your name here}`，`type` and `name` are required, `default` and `note` are optional. *It seems that `default` and `note` are not useful, but if you use xlreport as a ImagePy plugin, `default` and `note` would be parsed as some ui element.*

**Choices Parameter**

`list`:`{list Favourite_System = [Windows, Linux, Mac] # please select your favourite system}`，`default value` is required, should be a list. *If you use code to fill template, `choices` could be replaced by `str`. But if you sue xlreport as a ImagePy plugin, choices would be parsed as a combo list.*

**Image Parameter**

`img`:`{img My_Photo = [4.8,4.8,0.9,0] # you photo here`，`default value` is required, give 4 numbers，[`width` (in cm)，`height` (in cm)，`margin` (0.9 meas 10% margin)，`scaled` (0 means keep the proportion, 1 means scale to fit the cell's size），`note` is optional.

**Table Parameter**

`tab`:`{tab Record = [1,3,0,0] # the result table}`，`default value` is required, give 4 numbers, [`row step` (1 for unmerged cell), `column step` (1 for unmerged cell), `title offset` (-1 for the row upon the first data row, 0 means not fill title), `index offset` (-1 for column left of the first data column, 0 means not fill index)], `note` is optional.



## Use xlreport as ImagePy Plugin

[ImagePy](https://github.com/Image-Py/imagepy) is an Image Process Framework written in Python. I write xlreprot as ImagePy's report plugin, we can open image, do filter, analysis, and fill the image and table in template. You can also try [ImagePy](https://github.com/Image-Py/imagepy).

![](http://idoc.imagepy.org/demoplugin/37.png)

<div align=center>coins segment and measure</div><br>

![](http://idoc.imagepy.org/demoplugin/38.png)

<div align=center>generate report</div><br>