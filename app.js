const Express = require("express");
const Mongoose = require("mongoose");
const BodyParser = require("body-parser");
var app = Express();
app.use(BodyParser.json());
app.use(BodyParser.urlencoded({ extended: true }));
const excel = require('node-excel-export');
// Define REST API





Mongoose.connect("mongodb://localhost/demo",{ useNewUrlParser: true, useUnifiedTopology: true })
.then(() => console.log("Thành công"))
.catch(err => console.error(err));


const PersonModel = Mongoose.model("person", {
    firstname: String,
    lastname: String
});

app.post("/person", async (request, response) => {
    try {
        var person = new PersonModel(request.body);
        var result = await person.save();
        response.send(result);
    } catch (error) {
        response.status(500).send(error);
    }
});

app.post("/people", async (request, res) => {
    try {
        var result = await PersonModel.find().exec();
        const styles = {
            headerDark: {
            fill: {
                fgColor: {
                rgb: 'FF000000'
                }
            },
            font: {
                color: {
                rgb: 'FFFFFFFF'
                },
                sz: 14,
                bold: true,
                underline: true
            }
            },
            cellPink: {
            fill: {
                fgColor: {
                rgb: 'FFFFCCFF'
                }
            }
            },
            cellGreen: {
            fill: {
                fgColor: {
                rgb: 'FF00FF00'
                }
            }
            }
        };
        
       
        const heading = [
           
        ];
        
        const specification = {
            firstname: { 
                displayName: 'firstname', 
                headerStyle: styles.headerDark, 
                width: 120 
            },
            lastname: {
                displayName: 'lastname',
                headerStyle: styles.headerDark,
                width: '10'
            },
        }
        
      
        const report = excel.buildExport(
            [ 
            {
                name: 'Report', 
                heading: heading,
                specification: specification, 
                data: result 
            }
            ]
        );
        
        res.attachment('report.xlsx'); 
        res.send(report);
    } catch (error) {
        response.status(500).send(error);
    }
});
app.listen(3000, () => {
    console.log("Listening at :3000...");
});
