#pragma once

class Ado2Access
{
public:
	Ado2Access();
	_ConnectionPtr   m_pConnection;		// database
	_RecordsetPtr    m_pRecordset;		// command
	_CommandPtr      m_pCommand;		// record
	~Ado2Access();
};


