/*
 * Navigational Stars plug-in for Stellarium
 *
 * Copyright (C) 2016 Alexander Wolf
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program. If not, see <http://www.gnu.org/licenses/>.
 */

#include "NavStars.hpp"
#include "NavStarsWindow.hpp"
#include "ui_navStarsWindow.h"

#include "StelApp.hpp"
#include "StelCore.hpp"
#include "StelLocaleMgr.hpp"
#include "StelModule.hpp"
#include "StelModuleMgr.hpp"
#include "StelGui.hpp"
#include "StelUtils.hpp"
#include "StelObjectMgr.hpp"

#include <math.h>

#include <QComboBox>

#include <QFileDialog>
#include <QDir>
#include <QSortFilterProxyModel>
#include <QStringListModel>

#include "external/qcustomplot/qcustomplot.h"
#include "external/qxlsx/xlsxdocument.h"
#include "external/qxlsx/xlsxchartsheet.h"
#include "external/qxlsx/xlsxcellrange.h"
#include "external/qxlsx/xlsxchart.h"
#include "external/qxlsx/xlsxrichstring.h"
#include "external/qxlsx/xlsxworkbook.h"
using namespace QXlsx;


const QString NavStarsWindow::delimiter(", ");
#ifdef Q_OS_WIN
const QString NavStarsWindow::acEndl("\r\n");
#else
const QString NavStarsWindow::acEndl("\n");
#endif


NavStarsWindow::NavStarsWindow()
	: StelDialog("NavStars")
	, ns(Q_NULLPTR)
{
	ui = new Ui_navStarsWindowForm();
	predictionHeader.clear();
	core = StelApp::getInstance().getCore();
	objectMgr = GETSTELMODULE(StelObjectMgr);
}

NavStarsWindow::~NavStarsWindow()
{
	delete ui;
}

void NavStarsWindow::retranslate()
{
	if (dialog)
	{
		ui->retranslateUi(dialog);
		setAboutHtml();
		populateNavigationalStarsSets();
		populateNavigationalStarsSetDescription();
	}
}

void NavStarsWindow::createDialogContent()
{
	ns = GETSTELMODULE(NavStars);
	ui->setupUi(dialog);

	connect(&StelApp::getInstance(), SIGNAL(languageChanged()), this, SLOT(retranslate()));
	connect(ui->closeStelWindow, SIGNAL(clicked()), this, SLOT(close()));
	connect(ui->TitleBar, SIGNAL(movedTo(QPoint)), this, SLOT(handleMovedTo(QPoint)));

	populateNavigationalStarsSets();
	populateNavigationalStarsSetDescription();
	QString currentNsSetKey = ns->getCurrentNavigationalStarsSetKey();
	int idx = ui->nsSetComboBox->findData(currentNsSetKey, Qt::UserRole, Qt::MatchCaseSensitive);
	if (idx==-1)
	{
		// Use AngloAmerican as default
		idx = ui->nsSetComboBox->findData(QVariant("AngloAmerican"), Qt::UserRole, Qt::MatchCaseSensitive);
	}
	ui->nsSetComboBox->setCurrentIndex(idx);
	connect(ui->nsSetComboBox, SIGNAL(currentIndexChanged(int)), this, SLOT(setNavigationalStarsSet(int)));
	ui->displayAtStartupCheckBox->setChecked(ns->getEnableAtStartup());
	connect(ui->displayAtStartupCheckBox, SIGNAL(stateChanged(int)), this, SLOT(setDisplayAtStartupEnabled(int)));
	ui->displayDataOnScreenCheckBox->setChecked(ns->getEnableShowOnScreen());
	connect(ui->displayDataOnScreenCheckBox, SIGNAL(stateChanged(int)), this, SLOT(setDisplayDataOnScreenEnabled(int)));

	connect(ui->pushButtonSave, SIGNAL(clicked()), this, SLOT(saveSettings()));	
	connect(ui->pushButtonReset, SIGNAL(clicked()), this, SLOT(resetSettings()));

	// Twilight tab
	populateTwilightWindow();
	connect(ui->pushButtonRefreshTwilight, SIGNAL(clicked()), this, SLOT(populateTwilightWindow()));
	connect(ui->pushButtonRefreshPredictionList, SIGNAL(clicked()), this, SLOT(populateTwilightWindow()));
	connect(ui->pushButtonRefresh_3, SIGNAL(clicked()), this, SLOT(populateTwilightWindow()));
	// Since this automatiec update in time (hours, minutes or seconds) seems not to work properly,
	// a refresh-button was introduced and the automatic procedure was disabled.
//	connect(core, SIGNAL(locationChanged(StelLocation)), this, SLOT(populateTwilightWindow()));
//	connect(core, SIGNAL(dateChanged()), this, SLOT(populateTwilightWindow()));

	// Prediction tab
	populatePredictionList();
	connect(ui->pushButtonRefreshPredictionList, SIGNAL(clicked()), this, SLOT(populatePredictionList()));
	// This takes quite a long time so it was disabled
//	connect(ui->radioButtonMorningTwilight, SIGNAL(clicked()), this, SLOT(populatePredictionList()));
//	connect(ui->radioButtonEveningTwilight, SIGNAL(clicked()), this, SLOT(populatePredictionList()));
	connect(ui->nsPushButtonWriteToFile, SIGNAL(clicked()), this, SLOT(savePredictedPositions()));

	// About tab
	setAboutHtml();
	StelGui* gui = dynamic_cast<StelGui*>(StelApp::getInstance().getGui());
	if(gui!=Q_NULLPTR)
		ui->aboutTextBrowser->document()->setDefaultStyleSheet(QString(gui->getStelStyle().htmlStyleSheet));
}

void NavStarsWindow::populateTwilightWindow(void)
{
	//int jd = round(core->getJDE());	// round JD to next noon in Greenwich to get the best approximation for morning and evening
	int year, month, day, dayOfWeek(0);
	StelUtils::getDateFromJulianDay(core->getJDE(), &year, &month, &day);
	ui->labelDate  ->setText(QString("%1").arg(StelUtils::localeDateString(year, month, day, dayOfWeek)));	// TODO ist StelUtils hier noetig ?!?
	ui->labelDate_2->setText(QString("%1").arg(StelUtils::localeDateString(year, month, day, dayOfWeek)));
	ui->labelDate_3->setText(QString("%1").arg(StelUtils::localeDateString(year, month, day, dayOfWeek)));

//	double eot = core->getSolutionEquationOfTime(jd);


	double lon = core->getCurrentLocation().longitude;
	double lat = core->getCurrentLocation().latitude;
	ui->labelLatitude  ->setText(QString("%1").arg(StelUtils::radToDdmPStr(lat * M_PI_180f, 1, false, 'V')));	// TODO ist StelUtils hier noetig ?!?
	ui->labelLatitude_2->setText(QString("%1").arg(StelUtils::radToDdmPStr(lat * M_PI_180f, 1, false, 'V')));
	ui->labelLatitude_3->setText(QString("%1").arg(StelUtils::radToDdmPStr(lat * M_PI_180f, 1, false, 'V')));
	ui->labelLongitude  ->setText(QString("%1").arg(StelUtils::radToDdmPStr(lon * M_PI_180f, 1, false, 'H')));
	ui->labelLongitude_2->setText(QString("%1").arg(StelUtils::radToDdmPStr(lon * M_PI_180f, 1, false, 'H')));
	ui->labelLongitude_3->setText(QString("%1").arg(StelUtils::radToDdmPStr(lon * M_PI_180f, 1, false, 'H')));

	QString MorningTwilight, EveningTwilight, DurationTwilight, Transit;

	/* civil twilight */
	/* min altitude = - 0.83333 deg, this standard-value takes the refraction into account.
	 * Delta altitude = 6 deg - 0.83333 deg, mean altitude = min altitude - Delta altitude/2 */
	double angularSize = objectMgr->searchByName("Sun")->getAngularSize(core);
	computeTwilight(-3.41666667, maxDurationTwilight(-3.41666667, 5.16666), MorningTwilight, EveningTwilight, DurationTwilight, Transit);
	ui->labelBeginCivilTwilight->setText(MorningTwilight);
	ui->labelEndCivilTwilight->setText(EveningTwilight);
	ui->labelEveningTwilight->setText(EveningTwilight);

	/* nautical twilight */
	/* Delta altitude = 6 deg, mean altitude = -9 deg */
	computeTwilight(-9., maxDurationTwilight(-9., 6.), MorningTwilight, EveningTwilight, DurationTwilight, Transit);
	ui->labelBeginNavTwilight->setText(MorningTwilight);
	ui->labelEndNavTwilight->setText(EveningTwilight);
	ui->labelMorningTwilight->setText(MorningTwilight);
	ui->labelDuration->setText(DurationTwilight);

	/* astronomical twilight */
	/* Delta altitude = 6 deg, mean altitude = -15 deg */
	computeTwilight(-15., maxDurationTwilight(-15., 6.), MorningTwilight, EveningTwilight, DurationTwilight, Transit);
	ui->labelBeginAstroTwilight->setText(MorningTwilight);
	ui->labelEndAstroTwilight->setText(EveningTwilight);

	/* sunrise, sunset, and transit */
	/* This is a hack to reuse the function "computeTwilight()":
	 * 1. Set "alt" to the negative radius of the Sun
	 * 2. Set twilight duration to zero by multiplying the duration with zero, i.e. factor = 0 */
	computeTwilight(-angularSize, 0., MorningTwilight, EveningTwilight, DurationTwilight, Transit);
	ui->labelSunRise->setText(MorningTwilight);
	ui->labelSunSet->setText(EveningTwilight);
	ui->labelSunTransit->setText(Transit);

	// TODO Moon
//	computeTwilight(-angularSize, 0., MorningTwilight, EveningTwilight, DurationTwilight, Transit);
//	ui->labelMoonRise->setText(MorningTwilight);
//	ui->labelMoonSet->setText(EveningTwilight);
//	ui->labelMoonTransit->setText(Transit);

	// computeRiseSetTransit(0.); does not work currently
}

double NavStarsWindow::maxDurationTwilight(double meanAltitude, double DeltaAltitude)
{
	return 1440./360. * std::cos(meanAltitude*M_PI_180f) * DeltaAltitude;	// in minutes
}

void NavStarsWindow::computeRiseSetTransit(double alt)
{

	double storeJD = StelApp::getInstance().getCore()->getJD();
	qDebug() << "NJ: getJD " << storeJD;

//	StelApp::getInstance().getCore()->setJD(12345.6789);

	double lon = core->getCurrentLocation().longitude;

	double LHA = 1;				// inital value
	double JD = round(core->getJDE());	// inital value

	Vec3d pos;
	double ra, dec, dLHA;
	double r = 1;

	for (int i=0; i<10; i++){

		pos = objectMgr->searchByName("Moon")->getEquinoxEquatorialPos(core);
		StelUtils::rectToSphe(&ra, &dec, pos);

		qDebug() << "ra = " << ra << " dec = " << dec;

		r = std::cos(LHA) * std::cos(lon) * std::cos(dec) - std::sin(alt) + std::sin(lon) * std::sin(dec);

		qDebug() << "NJ: r = " << r;

		if (qAbs(r) < 1.e-4)
			break;

		dLHA = -r/(std::cos(lon) * std::cos(dec) * std::sin(LHA));
		LHA += dLHA;
		qDebug() << "i = " << i << " dLHA = " << dLHA << " LHA = " << LHA << " JD = " << JD;

		JD += dLHA/15.04107 / 24;
		StelApp::getInstance().getCore()->setJD(JD);
		qDebug() << "i = " << i << " JD = " << JD;
		qDebug() ;
	}

	StelApp::getInstance().getCore()->setJD(storeJD);
	qDebug() << "NJ: LHA = " << LHA;
}

//void NavStarsWindow::computeRiseSetTransit(double alt)
//{


//	double lon = core->getCurrentLocation().longitude;

//	double LHA = 0;				// inital value
//	double jd = round(core->getJDE());	// inital value

//	Vec3d pos;
//	double ra, dec, r, dLHA;

//	for (int i=0; i<10; i++){

//		pos = objectMgr->searchByName("Moon")->getEquinoxEquatorialPos(core);
//		StelUtils::rectToSphe(&ra, &dec, pos);


//		r = std::cos(LHA) * std::cos(lon) * std::cos(dec) - std::sin(alt) + std::sin(lon) * std::sin(dec);

//		if (abs(r) < 1.e-4)
//			break;

//		dLHA = std::cos(lon) * std::cos(dec) * std::sin(LHA);
//		LHA -= r/dLHA;


//		jd += dLHA/15.04107 / 24;
//		TODO Update jd;
//	}

//	return LHA;
//}

Vec3d NavStarsWindow::setEquitorialPos(QString CelestialObject)
{
	bool J2000EquPos = false;

	Vec3d pos;

	if (J2000EquPos)
	{
		pos = objectMgr->searchByName(CelestialObject)->getJ2000EquatorialPos(core);
	}
	else
	{
		pos = objectMgr->searchByName(CelestialObject)->getEquinoxEquatorialPos(core);
	}

	return pos;
}

void NavStarsWindow::computeTwilight(double alt, double maxDurationTwilight, QString &MorningTwilight,
				     QString &EveningTwilight, QString &DurationTwilight,
				     QString &Transit)
{
	// alt in degrees
	// rise, set, transit in UTC(hours), duration in min

	// Approximate calculation of begin, end, and duration of twilight in UTC

	/* Note: The begin and end of the twilight is a function of the
	 * declination and the equation of time. Both are not constant within
	 * a day but they vary only by a small amount. Therefore we compute the
	 * begin/end and the duration of the twilight by means of linearization
	 * of the governing equations. For that, we use the mean altitude (e.g.
	 * -9 degree for nautical twilight) instead of the standard definition
	 * of twilight (altitude = -6 ... -12 degrees for naut tw). The typical
	 * error of this approach is expected to lie in the range of some seconds
	 * which can sum up to one or two minutes due to rounding effects. The
	 * error is considered to be small compared to the potential error in
	 * refraction for bodies close to the horizon. Moreover there is no
	 * necessity to know the begin/end of the twilight by fractions of seconds
	 * in case of astronomical navigation purposes. */

	int jd = round(core->getJDE());	// round JD to next noon in Greenwich to get the best approximation for morning and evening
	double eot = core->getSolutionEquationOfTime(jd);
	double lon = core->getCurrentLocation().longitude;	//lon is only needed in degrees below and thus no transformation was done here
	double lat = (core->getCurrentLocation().latitude) * M_PI_180f;
	double ra, dec;
	StelUtils::rectToSphe(&ra, &dec, setEquitorialPos("Sun"));
	alt = alt * M_PI_180f;

	double arg = (std::sin(alt)-std::sin(lat)*std::sin(dec))/std::cos(lat)/std::cos(dec);

	if (qAbs(arg) > 1.) // no sunrise/sunset today at this latitude
	{
		if (arg > 1.) // The Sun remains below the horizon.
		{
			MorningTwilight = QString("--");
			EveningTwilight = QString("--");
			DurationTwilight = QString("--");
		}
		else // (arg < -1.) // The Sun is permanentely visible.
		{
			MorningTwilight = QString("--");
			EveningTwilight = QString("--");
			DurationTwilight = QString("--");
		}
	}
	else
	{
		double lhaAtSunSet = std::acos(arg);		// appox. LHA in rad at desired altitude (sunrise=360-sunset)

		double duration = maxDurationTwilight/std::cos(lat)/std::cos(dec)/std::sin(lhaAtSunSet); // in minutes
		DurationTwilight = QString("%1 min").arg((int)duration);

		double timeRise = 12. - eot/60. - lon/15. - lhaAtSunSet * M_180_PIf/15.; // UTC in hours
		timeRise -= duration/60. /2.;
		timeRise = fmod(timeRise + 48., 24.);		// force it to be in the range 0...24
		MorningTwilight = transformTimeToString(timeRise);

		double timeSet  = 12. - eot/60. - lon/15. + lhaAtSunSet * M_180_PIf/15.; // UTC in hours
		timeSet += duration/60. /2.;
		timeSet = fmod(timeSet + 48., 24.);		// force it to be in the range 0...24
		EveningTwilight = transformTimeToString(timeSet);
	}

	double transit = 12. - eot/60. - lon/15.;
	transit = fmod(transit + 48., 24.);
	Transit = transformTimeToString(transit);
}

QString NavStarsWindow::transformTimeToString(double time)	// siehe StelUtils ?!?
{
	QString str;
	QTextStream os(&str);

	int Hour = static_cast<int>(time);
	int Minute = (time - static_cast<float>(Hour)) * 60.;

	os << fixed;
	os << qSetFieldWidth(2) << qSetPadChar('0') << Hour;
	os << qSetFieldWidth(1) << ":";
	os << qSetFieldWidth(2) << qSetPadChar('0') << Minute << " UTC";

	return str;
}

void NavStarsWindow::setPredictionListHeaderNames()
{
	predictionHeader.clear();
	// TRANSLATORS: number of the Star in Nautical Almanac
	predictionHeader << q_("Star number");
	// TRANSLATORS: name of star
	predictionHeader << q_("Name");
	// TRANSLATORS: azimuth
	predictionHeader << q_("Azimuth");
	// TRANSLATORS: altitude
	predictionHeader << q_("Altitude");
	// TRANSLATORS: magnitude
	predictionHeader << q_("Mag.");
//	// TRANSLATORS: right ascension
//	predictionHeader << q_("RA (J2000)");
	// TRANSLATORS: declination
	predictionHeader << q_("Dec (J2000)");
	ui->predictionTreeWidget->setHeaderLabels(predictionHeader);

	// adjust the column width
	for (int i = 0; i < PredictionCount; ++i)
	{
		ui->predictionTreeWidget->resizeColumnToContents(i);
	}
}

void NavStarsWindow::initPredictionList()
{
	ui->predictionTreeWidget->clear();
	ui->predictionTreeWidget->setColumnCount(PredictionCount);
	setPredictionListHeaderNames();
	ui->predictionTreeWidget->header()->setSectionsMovable(false);
	ui->predictionTreeWidget->header()->setDefaultAlignment(Qt::AlignHCenter);
}

void NavStarsWindow::populatePredictionListLine(QString StarNumber, QString label, QString englishName,
						double magnitude, double azimuth, double altitude, double ra, double dec)
{
	NSPredictionTreeWidgetItem* treeItem = new NSPredictionTreeWidgetItem(ui->predictionTreeWidget);

	treeItem->setText(PredictionCONumber, StarNumber);

	treeItem->setText(PredictionCOName, label);
	treeItem->setData(PredictionCOName, Qt::UserRole, englishName);			// TODO what is englishName?

	treeItem->setText(PredictionMagnitude, QString::number(magnitude, 'f', 0));
	treeItem->setTextAlignment(PredictionMagnitude, Qt::AlignRight);

	treeItem->setText(PredictionAltitude, QString::number(altitude * M_180_PIf, 'f', 2));	// TODO in decimal degrees
	treeItem->setTextAlignment(PredictionAltitude, Qt::AlignRight);

	// TODO pruefe, wie sich das Azimut veraendert, wenn Einstellung sued-/nordazimut geschalten wird.
	treeItem->setText(PredictionAzimuth, QString::number(fmod( 180. - (azimuth * M_180_PIf) + 720., 360.), 'f', 2));		// TODO in decimal degrees
	treeItem->setTextAlignment(PredictionAzimuth, Qt::AlignRight);

	//treeItem->setText(PredictionRa, StelUtils::radToHmsStrAdapt(ra));		// TODO replace by star angle
	//treeItem->setTextAlignment(PredictionRa, Qt::AlignRight);

	treeItem->setText(PredictionDec, StelUtils::radToDdmPStr(dec, 1, false, 'V'));
	treeItem->setTextAlignment(PredictionDec, Qt::AlignRight);
}

void NavStarsWindow::setVisibleStarsAltitudeLimit(double &minAltitude, double &maxAltitude)
{
	// TODO prüfen und als Schalter einbauen, Grenzen beachten
	bool allVisible = false;

	minAltitude = 20 * M_PI_180f;
	maxAltitude = 90 * M_PI_180f;

	allVisible = true;		// TODO only for debugging.
	if (allVisible)
	{
		minAltitude =-90 * M_PI_180f;
		maxAltitude = 90 * M_PI_180f;
	}
}

void NavStarsWindow::populatePredictionList(void)
{
	double currentJD = core->getJD();   // save current JD
	//double currentJDE = core->getJDE(); // save current JDE

	double noonJD = round(core->getJD());	// round JD to next noon in Greenwich to get the best approximation for morning and evening
	core->setJD(noonJD);
	qDebug() << "NJ noonJD = " << fixed << qSetRealNumberPrecision(2) << noonJD;

	double lon = core->getCurrentLocation().longitude;
	double lat = core->getCurrentLocation().latitude;
	double eot = core->getSolutionEquationOfTime(noonJD);

	double ra, dec;
	StelUtils::rectToSphe(&ra, &dec, setEquitorialPos("Sun"));
	dec = dec * M_180_PIf;

	lat = lat * M_PI_180f;
	//lon is only needed in degrees below and thus no transformation was done here
	dec = dec * M_PI_180f;
	double alt = -9. * M_PI_180f;	// TODO stimmt der Winkel so?

	double arg = (std::sin(alt)-std::sin(lat)*std::sin(dec))/std::cos(lat)/std::cos(dec);

	double lhaAtSunSet = 0;

	if (qAbs(arg) > 1.) // no sunrise/sunset today at this latitude
	{
		// TODO
	}
	else
	{
		lhaAtSunSet = std::acos(arg);		// appox. LHA in rad at altitude (sunrise=360-sunset)
	}

	//double duration = maxDurationTwilight(-9., 6.)/std::cos(lat)/std::cos(dec)/std::sin(lhaAtSunSet); // in minutes

	double timeRise = 12. - eot/60. - lon/15. - lhaAtSunSet * M_180_PIf/15.; // UTC in hours

	double twiligthJD ;

	bool computeMorningTwiligth = true;
	if (computeMorningTwiligth)
	{
		//timeRise = timeRise - duration/60. /2.;	// in hours
		twiligthJD = noonJD - timeRise/24;
		qDebug() << "compute morning twilight";
	}
	else
	{
		//timeRise = timeRise + duration/60. /2.;
		twiligthJD = noonJD + timeRise/24;
		qDebug() << "compute evening twilight";
	}

	//timeRise = fmod(timeRise + 48., 24.);		// force it to be in the range 0...24
	//double twiligthJD = noonJD + timeRise/24.;

	twiligthJD = 11111111.1; // TODO ?!?!?!?!? warum hat das keinen Einfluss auf die Tabellenwerte?

	core->setJD(twiligthJD);
	qDebug() << "NJ twiligthJD = " << fixed << qSetRealNumberPrecision(2) << twiligthJD;





	initPredictionList();		// reset widget

	QList<int> starNumbers=ns->getStarsNumbers();

	QString name, englishName, label;
	double magnitude, altitude, azimuth; //, ra, dec;

	double minAltitude, maxAltitude;
	setVisibleStarsAltitudeLimit(minAltitude, maxAltitude);

	QVector<QString> visibleObject(6);
	visibleObject[0] = "Sun";
	visibleObject[1] = "Moon";
	visibleObject[2] = "Venus";
	visibleObject[3] = "Mars";
	visibleObject[4] = "Jupiter";
	visibleObject[5] = "Saturn";

	QVector<QString> visibleObjectSymbol(6);
	visibleObjectSymbol[0] = QChar(0x2609);
	visibleObjectSymbol[1] = QChar(0x263E);
	visibleObjectSymbol[2] = QChar(0x2640);
	visibleObjectSymbol[3] = QChar(0x2642);
	visibleObjectSymbol[4] = QChar(0x2643);
	visibleObjectSymbol[5] = QChar(0x2644);

	for (int i =0; i < visibleObject.size(); ++i)
	{
		name = visibleObject[i];

		label = objectMgr->searchByName(name)->getNameI18n();
		magnitude = objectMgr->searchByName(name)->getVMagnitude(core);

		StelUtils::rectToSphe(&azimuth, &altitude, objectMgr->searchByName(name)->StelObject::getAltAzPosAuto(core));
		StelUtils::rectToSphe(&ra, &dec, setEquitorialPos(name));

		if (altitude > minAltitude && altitude < maxAltitude)
		{
			populatePredictionListLine(visibleObjectSymbol[i], label, englishName, magnitude, azimuth, altitude, ra, dec);
		}
	}

	for (int i = 0; i < starNumbers.size(); ++i)
	{
		name = QString("HIP %1").arg(starNumbers.at(i));

		label = objectMgr->searchByName(name)->getNameI18n();
		magnitude = objectMgr->searchByName(name)->getVMagnitude(core);

		StelUtils::rectToSphe(&azimuth, &altitude, objectMgr->searchByName(name)->StelObject::getAltAzPosAuto(core));
		StelUtils::rectToSphe(&ra, &dec, setEquitorialPos(name));

		if (altitude > minAltitude && altitude < maxAltitude)
		{
			populatePredictionListLine(QString("(%1)").arg(i+1), label, englishName, magnitude, azimuth, altitude, ra, dec);
		}
	}

	for (int i = 0; i < PredictionCount; ++i)
	{
		ui->predictionTreeWidget->resizeColumnToContents(i);
	}

	core->setJD(currentJD); // restore time
	core->update(0);
}

void NavStarsWindow::cleanupPredictionList()
{
	ui->predictionTreeWidget->clear();
}

void NavStarsWindow::setAboutHtml(void)
{
	// Regexp to replace {text} with an HTML link.
	QRegExp a_rx = QRegExp("[{]([^{]*)[}]");

	QString html = "<html><head></head><body>";
	html += "<h2>" + q_("Navigational Stars Plug-in") + "</h2><table width=\"90%\">";
	html += "<tr width=\"30%\"><td><strong>" + q_("Version") + ":</strong></td><td>" + NAVSTARS_PLUGIN_VERSION + "</td></tr>";
	html += "<tr><td><strong>" + q_("License") + ":</strong></td><td>" + NAVSTARS_PLUGIN_LICENSE + "</td></tr>";
	html += "<tr><td><strong>" + q_("Author") + ":</strong></td><td>Alexander Wolf &lt;alex.v.wolf@gmail.com&gt;</td></tr>";
	html += "</table>";

	html += "<p>" + q_("This plugin marks navigational stars from a selected set.");
	html += "<p>";

	html += "<h3>" + q_("Links") + "</h3>";
	html += "<p>" + QString(q_("Support is provided via the Github website.  Be sure to put \"%1\" in the subject when posting.")).arg("Navigational Stars plugin") + "</p>";
	html += "<p><ul>";
	// TRANSLATORS: The text between braces is the text of an HTML link.
	html += "<li>" + q_("If you have a question, you can {get an answer here}.").toHtmlEscaped().replace(a_rx, "<a href=\"https://groups.google.com/forum/#!forum/stellarium\">\\1</a>") + "</li>";
	// TRANSLATORS: The text between braces is the text of an HTML link.
	html += "<li>" + q_("Bug reports and feature requests can be made {here}.").toHtmlEscaped().replace(a_rx, "<a href=\"https://github.com/Stellarium/stellarium/issues\">\\1</a>") + "</li>";
	// TRANSLATORS: The text between braces is the text of an HTML link.
	html += "<li>" + q_("If you want to read full information about this plugin and its history, you can {get info here}.").toHtmlEscaped().replace(a_rx, "<a href=\"http://stellarium.sourceforge.net/wiki/index.php/Navigational_Stars_plugin\">\\1</a>") + "</li>";
	html += "</ul></p></body></html>";


	StelGui* gui = dynamic_cast<StelGui*>(StelApp::getInstance().getGui());
	if(gui!=Q_NULLPTR)
	{
		QString htmlStyleSheet(gui->getStelStyle().htmlStyleSheet);
		ui->aboutTextBrowser->document()->setDefaultStyleSheet(htmlStyleSheet);
	}

	ui->aboutTextBrowser->setHtml(html);
}

void NavStarsWindow::saveSettings()
{
	ns->saveConfiguration();
}

void NavStarsWindow::resetSettings()
{
	ns->restoreDefaultConfiguration();
}

void NavStarsWindow::setDisplayAtStartupEnabled(int checkState)
{
	bool b = checkState != Qt::Unchecked;
	ns->setEnableAtStartup(b);
}

void NavStarsWindow::setDisplayDataOnScreenEnabled(int checkState)
{
	bool b = checkState != Qt::Unchecked;
	ns->setEnableShowOnScreen(b);
}

void NavStarsWindow::populateNavigationalStarsSets()
{
	Q_ASSERT(ui->nsSetComboBox);

	QComboBox* nsSets = ui->nsSetComboBox;

	//Save the current selection to be restored later
	nsSets->blockSignals(true);
	int index = nsSets->currentIndex();
	QVariant selectedNsSetId = nsSets->itemData(index);
	nsSets->clear();

	// TRANSLATORS: Part of full phrase: Anglo-American set of navigational stars
	nsSets->addItem(q_("Anglo-American"), "AngloAmerican");
	// TRANSLATORS: Part of full phrase: French set of navigational stars
	nsSets->addItem(q_("French"), "French");
	// TRANSLATORS: Part of full phrase: German set of navigational stars
	nsSets->addItem(q_("German"), "German");
	// TRANSLATORS: Part of full phrase: Russian set of navigational stars
	nsSets->addItem(q_("Russian"), "Russian");

	//Restore the selection
	index = nsSets->findData(selectedNsSetId, Qt::UserRole, Qt::MatchCaseSensitive);
	nsSets->setCurrentIndex(index);
	nsSets->blockSignals(false);
}

void NavStarsWindow::setNavigationalStarsSet(int nsSetID)
{
	QString currentNsSetID = ui->nsSetComboBox->itemData(nsSetID).toString();
	ns->setCurrentNavigationalStarsSetKey(currentNsSetID);
	populateNavigationalStarsSetDescription();
}

void NavStarsWindow::populateNavigationalStarsSetDescription(void)
{
	ui->nsSetDescription->setText(ns->getCurrentNavigationalStarsSetDescription());
}

void NavStarsWindow::savePredictedPositions()
{
	QString filter = q_("Microsoft Excel Open XML Spreadsheet");
	filter.append(" (*.xlsx);;");
	filter.append(q_("CSV (Comma delimited)"));
	filter.append(" (*.csv)");
	QString defaultFilter("(*.xlsx)");
	QString filePath = QFileDialog::getSaveFileName(Q_NULLPTR,
							q_("Save predicted positions of objects during twilight as..."),
							QDir::homePath() + "/PredictedPositions.xlsx",
							filter,
							&defaultFilter);

//	if (defaultFilter.contains(".csv", Qt::CaseInsensitive))
//		saveTableAsCSV(filePath, ui->predictionTreeWidget, predictionHeader);
//	else
	{
		int count = ui->predictionTreeWidget->topLevelItemCount();
		int columns = predictionHeader.size();

		int *width;
		width = new int[columns];
		QString sData;
		int w;

		QXlsx::Document xlsx;
		xlsx.setDocumentProperty("title", q_("Predicted positions of objects during twilight"));
		xlsx.setDocumentProperty("creator", StelUtils::getApplicationName());
		//xlsx.addSheet(ui->celestialCategoryComboBox->currentData(Qt::DisplayRole).toString(), AbstractSheet::ST_WorkSheet);

		QXlsx::Format header;
		header.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
		header.setPatternBackgroundColor(Qt::yellow);
		header.setBorderStyle(QXlsx::Format::BorderThin);
		header.setBorderColor(Qt::black);
		header.setFontBold(true);
		for (int i = 0; i < columns; i++)
		{
			// Row 1: Names of columns
			sData = predictionHeader.at(i).trimmed();
			xlsx.write(1, i + 1, sData, header);
			width[i] = sData.size();
		}

		QXlsx::Format data;
		data.setHorizontalAlignment(QXlsx::Format::AlignRight);
		for (int i = 0; i < count; i++)
		{
			for (int j = 0; j < columns; j++)
			{
				// Row 2 and next: the data
				sData = ui->predictionTreeWidget->topLevelItem(i)->text(j).trimmed();
				xlsx.write(i + 2, j + 1, sData, data);
				w = sData.size();
				if (w > width[j])
				{
					width[j] = w;
				}
			}
		}

		for (int i = 0; i < columns; i++)
		{
			xlsx.setColumnWidth(i+1, width[i]+2);
		}

		delete[] width;

//		// Add the date and time info for celestial positions
//		xlsx.write(count + 2, 1, ui->celestialPositionsTimeLabel->text());
//		QXlsx::CellRange range = CellRange(count+2, 1, count+2, columns);
//		QXlsx::Format extraData;
//		extraData.setBorderStyle(QXlsx::Format::BorderThin);
//		extraData.setBorderColor(Qt::black);
//		extraData.setPatternBackgroundColor(Qt::yellow);
//		extraData.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
//		xlsx.mergeCells(range, extraData);

		xlsx.saveAs(filePath);
	}
}


void NavStarsWindow::saveTableAsCSV(const QString &fileName, QTreeWidget* tWidget, QStringList &headers)
{
	int count = tWidget->topLevelItemCount();
	int columns = headers.size();

	QFile table(fileName);
	if (!table.open(QFile::WriteOnly | QFile::Truncate))
	{
		qWarning() << "NavStars: Unable to open file" << QDir::toNativeSeparators(fileName);
		return;
	}

	QTextStream tableData(&table);
	tableData.setCodec("UTF-8");

	for (int i = 0; i < columns; i++)
	{
		QString h = headers.at(i).trimmed();
		if (h.contains(","))
			tableData << QString("\"%1\"").arg(h);
		else
			tableData << h;

		if (i < columns - 1)
			tableData << delimiter;
		else
			tableData << acEndl;
	}

	for (int i = 0; i < count; i++)
	{
		for (int j = 0; j < columns; j++)
		{
			tableData << tWidget->topLevelItem(i)->text(j);
			if (j < columns - 1)
				tableData << delimiter;
			else
				tableData << acEndl;
		}
	}

	table.close();
}
