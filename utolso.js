const puppeteer = require('puppeteer');
const fs = require('fs');
const xlsx = require('xlsx');
const nodeMailer = require('nodemailer');

const app = async () => {
    const browser = await puppeteer.launch({ headless: true, 
        args: ['--no-sandbox', '--disable-setuid-sandbox']
     });
    const page = await browser.newPage();
    const d = new Date();
    const y = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = d.getDate();
    const date = `${y}-${month}-${day}`;

    var url = 'https://www.mbkandallo.hu/index.php?route=product%2Flist&keyword=samott';
    await page.goto(url, { waitUntil: 'networkidle2' });

    // lapszámok
    const pagesCount = await page.$$eval(
        '#mm-62 > main > div > div > section > div.page-body > div.sortbar.sortbar-bottom.section__spacer > div > nav > a', pages => (pages.length - 2)
    );

    console.log(`Found ${pagesCount} pages.`);

    const results = [];
    const dateHeader = date; // Excel header 

    for (let j = 1; j <= pagesCount; j++) {
        // browser megnyitas - query
        const pageLink = `https://www.mbkandallo.hu/index.php?route=product/list&keyword=Samott&page=${j}#content`;
        console.log(`Visiting: ${pageLink}`);
        await page.goto(pageLink, { waitUntil: 'networkidle2' });

        // product lista 
        const currentProductsCount = await page.$$eval(
            '#mm-62 > main > div > div > section > div.page-body > div.product-list.show-quantity-before-cart.section__spacer > div > div',
            products => products.length
        );

        console.log(`Found ${currentProductsCount} products on page ${j}`);

        // oldal loop
        for (let i = 1; i <= currentProductsCount; i++) {
            const productSelector = `#mm-62 > main > div > div > section > div.page-body > div.product-list.show-quantity-before-cart.section__spacer > div > div:nth-child(${i}) > div > div.card__footer.product-card__footer > div.product-card__item.product-card__details > a`;

            try {
                const productLink = await page.$eval(productSelector, el => el.href);
                console.log(`Visiting: ${productLink}`);
                await page.goto(productLink, { waitUntil: 'networkidle2' });

                // selectorok
                const availabilitySelector = '#product > div.product-page-top__row.row > div.col-lg-4.product-page-middle > div > table > tbody > tr.product-parameter.product-parameter__stock > td.product-parameter__value > span';
                const nameSelector = '#product > div.product-page-top__row.row > div.col-lg-4.product-page-middle > h1 > span';
                const typeSelector = '#product > div.product-page-top__row.row > div.col-lg-3.product-page-right > div.product-cart-box.d-flex.flex-column > div.product-attributes.product-page-right-box > div:nth-child(1) > span > ul > li:nth-child(2) > a';
                
                const vanetype = '#product-image > div.product-image__main > div.product_badges.horizontal-orientation > div.badgeitem-content.badgeitem-content-id-26.badgeitem-content-color-z.badgeitem-content-type-7 > a > span';
                const vanesize = '#product-image > div.product-image__main > div.product_badges.horizontal-orientation > div.badgeitem-content.badgeitem-content-id-22.badgeitem-content-color-c.badgeitem-content-type-7 > a > span';

                await page.waitForSelector(availabilitySelector, { timeout: 2500 });
                const availability = await page.$eval(availabilitySelector, el => el.textContent.trim());
                await page.waitForSelector(nameSelector, {timeout: 2500});
                const name = await page.$eval(nameSelector, el => el.textContent.trim());

                // result push + meret es tipus loop
                results.push({ name, availability });
                console.log(`Availability for ${name}: ${availability}`);
                const type = await page.$(vanetype);
                const meret = await page.$(vanesize);
                if(meret) {
                    console.log('VAN MERET');
                    const sizeListSelector = '#product > div.product-page-top__row.row > div.col-lg-3.product-page-right > div.product-cart-box.d-flex.flex-column > div.product-attributes.product-page-right-box > div:nth-child(2) > span > ul';
                    const currentSizeCount = await page.$$eval(`${sizeListSelector} > li`, sizes => sizes.length);
                    console.log(currentSizeCount);
                    for(let k = 2; k <= currentSizeCount; k++) {
                        const sizeSelector = `#product > div.product-page-top__row.row > div.col-lg-3.product-page-right > div.product-cart-box.d-flex.flex-column > div.product-attributes.product-page-right-box > div:nth-child(2) > span > ul > li:nth-child(${k}) > a`;
                        const sizeLink = await page.$eval(sizeSelector, el => el.href);
                        console.log(`Visiting: ${sizeLink}`);
                        await page.goto(sizeLink, {waitUntil: 'networkidle2'});
                        await page.waitForSelector(availabilitySelector, { timeout: 2500 });
                        const availability = await page.$eval(availabilitySelector, el => el.textContent.trim());
                        await page.waitForSelector(nameSelector, {timeout: 2500});
                        const name = await page.$eval(nameSelector, el => el.textContent.trim());
                        results.push({name, availability});
                        console.log(`Availability for ${name}: ${availability}`);
                    }
                    if(type) {
                        console.log('VAN TIPUS');
                        const typeLink = await page.$eval(typeSelector, el => el.href);
                        console.log(`Visiting: ${typeLink}`);
                        await page.goto(typeLink, {waitUntil: 'networkidle2'});
                        await page.waitForSelector(availabilitySelector, { timeout: 2500 });
                        const availability = await page.$eval(availabilitySelector, el => el.textContent.trim());
                        await page.waitForSelector(nameSelector, {timeout: 2500});
                        const name = await page.$eval(nameSelector, el => el.textContent.trim());
                        results.push({name, availability});
                        console.log(`Availability for ${name}: ${availability}`);
                        for(let k = 2; k <= currentSizeCount; k++) {
                            const sizeSelector = `#product > div.product-page-top__row.row > div.col-lg-3.product-page-right > div.product-cart-box.d-flex.flex-column > div.product-attributes.product-page-right-box > div:nth-child(2) > span > ul > li:nth-child(${k}) > a`;
                            const sizeLink = await page.$eval(sizeSelector, el => el.href);
                            console.log(`Visiting: ${sizeLink}`);
                            await page.goto(sizeLink, {waitUntil: 'networkidle2'});
                            await page.waitForSelector(availabilitySelector, { timeout: 2500 });
                            const availability = await page.$eval(availabilitySelector, el => el.textContent.trim());
                            await page.waitForSelector(nameSelector, {timeout: 2500});
                            const name = await page.$eval(nameSelector, el => el.textContent.trim());
                            results.push({name, availability});
                            console.log(`Availability for ${name}: ${availability}`);
                        }
                    }
                }
                if(type) {
                    console.log('VAN TIPUS');
                    const typeLink = await page.$eval(typeSelector, el => el.href);
                    console.log(`Visiting: ${typeLink}`);
                    await page.goto(typeLink, {waitUntil: 'networkidle2'});
                    await page.waitForSelector(availabilitySelector, { timeout: 2500 });
                            const availability = await page.$eval(availabilitySelector, el => el.textContent.trim());
                            await page.waitForSelector(nameSelector, {timeout: 2500});
                            const name = await page.$eval(nameSelector, el => el.textContent.trim());
                            results.push({name, availability});
                            console.log(`Availability for ${name}: ${availability}`);
                }
                // back to page
                await page.goto(pageLink, { waitUntil: 'networkidle2' });

            } catch (error) {
                console.error(`${i} product error:`, error);
                

                // back to page but w error
                await page.goto(pageLink, { waitUntil: 'networkidle2' });
            }
        }
    }

    // sheet data preparin
    const sheetData = results.map(product => ({
        'név': product.name, // Change 'name' to 'név'
        [date]: product.availability
    }));

    // fajl ellenorzes
        let wb, ws, existingData;
        const fileName = 'samott.xlsx';
    
        if (fs.existsSync(fileName)) {
            // If the file exists, read it
            wb = xlsx.readFile(fileName);
            ws = wb.Sheets["Sheet1"];
            existingData = xlsx.utils.sheet_to_json(ws, { defval: "" });
        } else {
            // If the file doesn't exist, initialize an empty workbook and data array
            wb = xlsx.utils.book_new();
            existingData = [];
        }
    
        // Extract headers from existing data, or create a new header row
        const headers = existingData.length > 0 ? Object.keys(existingData[0]) : ["név"];
    
        // Add the current date as a new column if it doesn't exist
        if (!headers.includes(date)) {
            headers.push(date);
        }
    
        // Ensure all names are in the sheet
        results.forEach(product => {
            const existingRow = existingData.find(row => row["név"] === product.name);
            if (!existingRow) {
                // Add a new row for this product if it doesn't already exist
                const newRow = { "név": product.name };
                headers.forEach(header => {
                    if (!newRow[header]) {
                        newRow[header] = ""; // Initialize other columns with empty strings
                    }
                });
                existingData.push(newRow);
            }
        });
    
        // Update availability for each product
        results.forEach(product => {
            const row = existingData.find(row => row["név"] === product.name);
            if (row) {
                row[date] = product.availability; // Add availability under the current date column
            }
        });
    
        // Convert the updated data back into a sheet
        ws = xlsx.utils.json_to_sheet(existingData, { header: headers });
    
        // Dynamically adjust column widths
        const colWidths = headers.map(() => ({ wch: 20 })); // Adjust widths as needed
        ws["!cols"] = colWidths;
    
        // Remove the existing sheet if it exists
        const sheetName = "Sheet1";
        if (wb.SheetNames.includes(sheetName)) {
            delete wb.Sheets[sheetName];
            const index = wb.SheetNames.indexOf(sheetName);
            wb.SheetNames.splice(index, 1);
        }
    
        // Append the new sheet with the same name
        xlsx.utils.book_append_sheet(wb, ws, sheetName);
        xlsx.writeFile(wb, fileName);
    
        console.log(`Results saved to ${fileName}`);

    // email
   /* const transporter = nodeMailer.createTransport({
        host: 'smtp.gmail.com',
        port: 465,
        secure: true,
        auth: {
            user: 'szdgwebwork@gmail.com',
            pass: 'mmvq xcrw zwae yath'
        }
    });

    const info = await transporter.sendMail({
        from: 'szdgwebwork <szdgwebwork@gmail.com>',
        to: 'Dozsa.miklos01@gmail.com',
        subject: `Napi samott: ${date}`,
        text: 'Jó napot! Itt a napi jelentés!',
        attachments: [{
            filename: fileName,
            path: `./${fileName}`
        }]
    });

    console.log('Message sent!');
*/
    await browser.close();

    setTimeout(app, 24 * 60 * 60 * 1000);
};

app();