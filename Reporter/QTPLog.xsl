<xsl:transform version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:template match="/"> 
 <html>
 <body>
 <h1>ReporterManager XML logs</h1>
 <br/>
  <xsl:for-each select="QTPLogs/QTPLog">	
	<table border="4" bordercolor="DarkGray" width="100%" bgcolor="WhiteSmoke">
		<tr border="0">
			<td align="center" bgcolor="#0066FF"><b>Log #<xsl:number value ="position()"/></b></td>			
		</tr>
		<tr>
			<td>
				<table border="0" width="90%" align="center">
					<tr bgcolor="Silver">
						<th width="25%">Test Name</th>
						<th width="25%">Status</th>
						<th width="25%">Started</th>
						<th width="25%">Ended</th>
					</tr>
					<tr>
						<td><xsl:value-of select="@TestName"/></td>
						
						<xsl:choose>
							<xsl:when test="@Status='Fail'">
								<td bgcolor="red" align="center"><xsl:value-of select="@Status"/></td>		
							</xsl:when>
							<xsl:when test="@Status='Warning'">
								<td bgcolor="gold" align="center"><xsl:value-of select="@Status"/></td>		
							</xsl:when>
							<xsl:when test="@Status='Pass'">
								<td bgcolor="green" align="center"><xsl:value-of select="@Status"/></td>		
							</xsl:when>
							<xsl:otherwise>
								<td align="center"><xsl:value-of select="@Status"/></td>		
							</xsl:otherwise>
						</xsl:choose>
						
						<td><xsl:value-of select="@Start"/></td>
						
						<td><xsl:value-of select="@End"/></td>
					</tr>
				</table>		
			</td>
		</tr>
		<tr>
			<td colspan="3">
				<table border="1" width="100%">
				<xsl:for-each select="Iteration">
					<tr>
						<td colspan="7" align="center" bgcolor="SkyBlue">
							<b>Iteration <xsl:value-of select="@Number"/>, Started <xsl:value-of select="@Start"/> - Status = 
								<xsl:choose>
									<xsl:when test="@Status='Fail'">
										<font color="red">Fail</font>
									</xsl:when>
									<xsl:when test="@Status='Warning'">
										<font color="gold">Warning</font>
									</xsl:when>
									<xsl:when test="@Status='Pass'">
										<font color="green">Pass</font>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="@Status"/>
									</xsl:otherwise>
								</xsl:choose>						
							</b>															
						</td>
					</tr>
					<tr bgcolor="Silver">
						<th width="5%">#</th>
						<th width="10%">Time</th>
						<th width="5%">Status</th>
						<th width="10%">Step Name</th>
						<th width="20%">Expected</th>
						<th width="20%">Actual</th>
						<th width="35%">Details</th>
					</tr>
					
					<xsl:for-each select="Step">
						<tr>
							<td align="center"><xsl:number value ="position()"/></td>
							<td><xsl:value-of select="Time"/></td>
							<td align="center">
								<xsl:choose>
									<xsl:when test="Status='Fail'">
										<font color="red">Fail</font>
									</xsl:when>
									<xsl:when test="Status='Warning'">
										<font color="gold">Warning</font>
									</xsl:when>
									<xsl:when test="Status='Pass'">
										<font color="green">Pass</font>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="Status"/>
									</xsl:otherwise>
								</xsl:choose>							
							</td>
							<td><xsl:value-of select="StepName"/></td>
							<td><xsl:value-of select="Expected"/></td>
							<td><xsl:value-of select="Actual"/></td>
							<td><xsl:value-of select="Details"/></td>
						</tr>
					</xsl:for-each><!-- Step -->
				
				</xsl:for-each> <!-- Iteration -->				
				</table>
			</td>
		</tr>	
		
	</table>
	<br/><br/>
  </xsl:for-each> <!-- Log -->
 </body>
 </html>


</xsl:template>
</xsl:transform>
