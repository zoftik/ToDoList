//required LB
const mongoose = require('mongoose');

//connect to DB
mongoose.connect('mongodb://127.0.0.1:27017/test');

// const Cat = mongoose.model('Cat', { name: String });

// const kitty = new Cat({ name: 'Zildjian' });
// kitty.save().then(() => console.log('meow'))


//acquire the connection(to check if successful)
const db = mongoose.connection;

//error
db.on ('error', console.error.bind(console, 'error connecting to db'));

//up and running then print
db.once('open', function(){
    console.log('succesfully connected to database');
});