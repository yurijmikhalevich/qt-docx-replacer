#ifndef DOCXREPLACER_H
#define DOCXREPLACER_H

#include <QString>
#include <QMap>

namespace DocxReplacer
{
bool replaceInFile(QString fileName, QMap<QString, QString> *replaceRules, QString newFileName = QString());
bool removeFolder(QString path);
}

#endif // DOCXREPLACER_H
