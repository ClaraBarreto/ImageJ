#@ ImagePlus imp
#@ ResultsTable rt

import ij.IJ
import fiji.analyze.directionality.Directionality_
import ij.WindowManager

d = new Directionality_()

d.setImagePlus(imp)
// d.setMethod(Directionality_.AnalysisMethod.FOURIER_COMPONENTS)
// d.setBinNumber(90)
d.computeHistograms()
d.fitHistograms()

result = d.getFitAnalysis()
imp = IJ.getImage()

mainDirection = Math.toDegrees(result[0][0])
dispersion = Math.toDegrees(result[0][1])
amount = result[0][2]
goodness = result[0][3]

// plot_frame = d.plotResults()
// plot_frame.show()
// plot_frame.hide()

rt.incrementCounter()
rt.addLabel(imp.getTitle())
rt.addValue("Main Direction", mainDirection)
rt.addValue("Dispersion", dispersion)
rt.addValue("Amount", amount)
rt.addValue("Goodness", goodness)
rt.show("Results")