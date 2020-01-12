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

#ifndef NAVSTARSWINDOW_HPP
#define NAVSTARSWINDOW_HPP

#include "StelDialog.hpp"
#include <QTreeWidgetItem>

class Ui_navStarsWindowForm;
class NavStars;

//! Main window of the Navigational Stars plug-in.
//! @ingroup navStars
class NavStarsWindow : public StelDialog
{
	Q_OBJECT

public:

	//! Defines the number and the order of the columns in the prediction table
	//! @enum PredictionColumns
	enum PredictionColumns {
		PredictionCONumber,		//! star number in chosen nautical almanac
		PredictionCOName,		//! name of navigational star
		PredictionAzimuth,		//! azimuth
		PredictionAltitude,		//! altitude
		PredictionMagnitude,		//! magnitude
		//PredictionRa,			//! Ra
		PredictionDec,			//! Dec
		PredictionCount			//! total number of columns
	};

	NavStarsWindow();
	~NavStarsWindow();

public slots:
	void retranslate();
	void cleanupPredictionList();
	void savePredictedPositions();

protected:
	void createDialogContent();

private:
	Ui_navStarsWindowForm* ui;
	NavStars* ns;
	class StelCore* core;
	class StelObjectMgr* objectMgr;

	void saveTableAsCSV(const QString &fileName, QTreeWidget* tWidget, QStringList &headers);

	void setAboutHtml();
	void populateNavigationalStarsSetDescription();
	double maxDurationTwilight(double meanAltitude, double);
	Vec3d setEquitorialPos(QString CelestialObject);
	void computeTwilight(double alt,
			     double maxDurationTwilight, QString &MorningTwilight,
			     QString &EveningTwilight, QString &duration, QString &Transit);
	QString transformTimeToString(double time);
	void computeRiseSetTransit(double alt);
	void setPredictionListHeaderNames();
	void initPredictionList();
	void setVisibleStarsAltitudeLimit(double &minAltitude, double &maxAltitude);

	QStringList predictionHeader;
	//! List of pointers to the objects representing the stars.
	QVector<StelObjectP> stars;

	static const QString delimiter, acEndl;

private slots:
	void saveSettings();
	void resetSettings();
	void setDisplayAtStartupEnabled(int checkState);
	void setDisplayDataOnScreenEnabled(int checkState);

	void populatePredictionListLine(QString StarNumber, QString label, QString englishName,
					double magnitude, double azimuth, double altitude,
					double ra, double dec);
	void populateNavigationalStarsSets();
	void setNavigationalStarsSet(int nsSetID);

	void populateTwilightWindow();
	void populatePredictionList();
};

// Reimplements the QTreeWidgetItem class to fix the sorting bug
class NSPredictionTreeWidgetItem : public QTreeWidgetItem
{
public:
	NSPredictionTreeWidgetItem(QTreeWidget* parent)
		: QTreeWidgetItem(parent)
	{
	}

private:
//	bool operator < (const QTreeWidgetItem &other) const
//	{
//		int column = treeWidget()->sortColumn();
//
//		if (column == NavStarsWindow::TransitDate)
//		{
//			return data(column, Qt::UserRole).toFloat() < other.data(column, Qt::UserRole).toFloat();
//		}
//		else if (column == NavStarsWindow::TransitMagnitude)
//		{
//			return text(column).toFloat() < other.text(column).toFloat();
//		}
//		else if (column == NavStarsWindow::TransitAltitude || column == AstroCalcDialog::TransitElongation || column == AstroCalcDialog::TransitAngularDistance)
//		{
//			return StelUtils::getDecAngle(text(column)) < StelUtils::getDecAngle(other.text(column));
//		}
//		else
//		{
//			return text(column).toLower() < other.text(column).toLower();
//		}
//	}
};

#endif /* NAVSTARSWINDOW_HPP */
