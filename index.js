#!/usr/bin/env node

'use strict';

const program = require('commander');

program
.version('0.0.1')
.description('Document generator from salesforce metadata');

require('./src/commands');


program.parse(process.argv)