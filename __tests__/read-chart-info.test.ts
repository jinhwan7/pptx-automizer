import Automizer, { read } from '../src/index';
import { vd } from '../src/helper/general-helper';

test('read chart info from workbook, e.g. series color', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartBarsStacked.pptx`, 'charts');

  const info = {
    series: [],
  };

  await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [read.readChartInfo(info)]);
    })
    .write(`read-chart-info.test.pptx`);

  console.log(info);
});
