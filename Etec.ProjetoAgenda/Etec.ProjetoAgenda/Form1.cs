using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
namespace Etec.ProjetoAgenda
{
    public partial class FAgenda : Form
    {
        static string strCn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Aluno_2\Desktop\BDAgenda1.accdb";
        OleDbConnection conexao = new OleDbConnection(strCn);
        public FAgenda()
        {
            InitializeComponent();
        }

        private void btnPesquisar_Click(object sender, EventArgs e)
        {
            if (txtId.Text == "")
            {
                MessageBox.Show("Valor Nulo! Digite um Id para pesquisar");
            }
            //instrução sql responsável por pesquisar o banco de dados (CRUD - Read)
            string pesquisa = "select * from tbcontato where Id = " + txtId.Text;

            //criando um objeto de nome cmd tendo como modelo a classe OleDbCommand para executar a instrução sql
            OleDbCommand cmd = new OleDbCommand(pesquisa, conexao);

            // Atravé da classe OleDbDataReader que faz parte do SqlCliente, criamos uma //variável chamada DR que será usada na leitura dos dados (instrução select)
            OleDbDataReader DR;

            //tratamento de exceções: try - catch - finally (em caso de erro capturamos o //tipo do erro)

            try
            {
                // Abrindo a conexão com o banco
                conexao.Open();
                // Executando a instrução e armazenando o resultado no reader DR
                DR = cmd.ExecuteReader();
                // Se houver um registro correspondente ao Id
                if (DR.Read())
                {
                    // Exibe as informações nas caixas de texto (textBox) correspondentes (0) //corresponde ao Id, (1) ao Nome e assim sucessivamente 
                    txtId.Text = DR.GetValue(0).ToString();
                    txtNome.Text = DR.GetValue(1).ToString();
                    txtFone.Text = DR.GetValue(2).ToString();
                    txtEmail.Text = DR.GetValue(3).ToString();
                }
                // Senão, exibimos uma mensagem avisando e também limpamos os campos para uma //nova pesquisa 
                else
                {
                    MessageBox.Show("Registro não encontrado");
                    txtNome.Clear();
                    txtFone.Clear();
                    txtEmail.Clear();
                    txtId.Focus();

                } // Encerrando o uso do reader DR 
                DR.Close();

                // Encerrando o uso do cmd 
                cmd.Dispose();
            }

                //caso ocorra algum erro 

            catch (Exception ex)
            {


            }

                                // de qualquer forma sempre fechar a conexão com o banco ("lembrar da porta da //geladeira rsrsrs") 
            finally
            {
                conexao.Close();
            }
        }

        private void btnAdicionar_Click(object sender, EventArgs e)
        {
            if (txtId.Text == "" || txtNome.Text == "" || txtFone.Text == "" || txtEmail.Text == "")
            {
                MessageBox.Show("Valores Nulos! Digite valores para adicionar");
            }
            else
            {
                //instrução sql responsável por adicionar dados ao banco (CRUD - Create) 
                string adiciona = "insert into tbcontato values (" +
                txtId.Text + ",'" +
                txtNome.Text + "','" +
                txtFone.Text + "','" +
                txtEmail.Text + "')";

                //criando um objeto de nome cmd tendo como modelo a classe OleDbCommand para //executar a instrução sql 
                OleDbCommand cmd = new OleDbCommand(adiciona, conexao);

                //tratamento de exceções: try - catch - finally (em caso de erro capturamos o //tipo do erro) 
                try
                {
                    // Abrindo a conexão com o banco 
                    conexao.Open();

                    // Criando uma variável para adicionar e armazenar o resultado 
                    int resultado;
                    resultado = cmd.ExecuteNonQuery();
                    // Verificando se o registro foi adicionado 
                    // Caso o valor da variável resultado seja 1 
                    // significa que o comando funcionou, neste caso limpar os campos e exibir uma //mensagem 
                    if (resultado == 1)
                    {
                        MessageBox.Show("Registro adicionado com sucesso");
                        txtId.Clear();
                        txtNome.Clear();
                        txtFone.Clear();
                        txtEmail.Clear();
                        txtId.Focus();
                    }
                    // Encerrando o uso do cmd 
                    cmd.Dispose();
                }

                        //caso ocorra algum erro 
                catch (Exception ex)
                {


                }

                            // de qualquer forma sempre fechar a conexão com o banco ("lembrar da porta da //geladeira rsrsrs") 
                finally
                {
                    conexao.Close();
                }


            }
        }
        private void btnAlterar_Click(object sender, EventArgs e)
        {
            if (txtId.Text == "")
            {
                MessageBox.Show("Valor Nulo! Digite um Id para Alterar");
            }
            //instrução sql responsável por alterar um registro do banco (CRUD - Update) 
            string altera = "update tbcontato set Nome= '" + txtNome.Text +
            "', Fone= '" + txtFone.Text +
            "', Email= '" + txtEmail.Text +
            "' where Id= " + txtId.Text;

            //criando um objeto de nome cmd tendo como modelo a classe OleDbCommand para //executar a instrução sql 
            OleDbCommand cmd = new OleDbCommand(altera, conexao);

            //tratamento de exceções: try - catch - finally (em caso de erro capturamos o //tipo do erro) 
            try
            {
                // Abrindo a conexão com o banco 
                conexao.Open();

                // Criando uma variável para alterar e armazenar o resultado 
                int resultado;
                resultado = cmd.ExecuteNonQuery();
                // Verificando se o registro foi alterado 
                // Caso o valor da variável resultado seja 1 
                // significa que o comando funcionou, neste caso limpar os campos e exibir uma //mensagem 
                if (resultado == 1)
                {
                    txtId.Clear();
                    txtNome.Clear();
                    txtFone.Clear();
                    txtEmail.Clear();
                    txtId.Focus();
                    MessageBox.Show("Registro alterado com sucesso");
                }
                // Encerrando o uso do cmd 
                cmd.Dispose();
            }

                    //caso ocorra algum erro 
            catch (Exception ex)
            {

            }
            // De qualquer forma sempre fechar a conexão com o banco 
            finally
            {
                conexao.Close();
            }

        }

        private void btnExcluir_Click(object sender, EventArgs e)
        {
            if (txtId.Text == "")
            {
                MessageBox.Show("Valor Nulo! Digite um Id para excluir");
            }
            //instrução sql responsável por remover um registro do banco (CRUD - Delete) 
            string remove = "delete from tbcontato where Id= " + txtId.Text;
            //criando um objeto de nome cmd tendo como modelo a classe OleDbCommand para //executar a instrução sql 
            //criando um objeto de nome cmd tendo como modelo a classe OleDbCommand para //executar a instrução sql 
            OleDbCommand cmd = new OleDbCommand(remove, conexao);
            //tratamento de exceções: try - catch - finally (em caso de erro capturamos o //tipo do erro) 
            try
            {

                // Abrindo a conexão com o banco 
                conexao.Open();
                // Criando uma variável para adicionar e armazenar o resultado 
                int resultado;
                if (txtId.Text != "")
                {
                    if (MessageBox.Show("Tem certeza que deseja remover este registro ?", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        resultado = cmd.ExecuteNonQuery();
                        // Verificando se o registro foi apagado 
                        // Caso o valor da variável resultado seja 1 
                        // significa que o comando funcionou, neste caso limpar os campos e exibir uma //mensagem 
                        if (resultado == 1)
                        {
                            txtId.Clear();
                            txtNome.Clear();
                            txtFone.Clear();
                            txtEmail.Clear();
                            txtId.Focus();
                            MessageBox.Show("Registro removido com sucesso");
                        }
                        // Encerrando o uso do cmd 
                        cmd.Dispose();
                    }
                }
            }
            //caso ocorra algum erro 
            catch (Exception ex)
            {

            }
            // de qualquer forma sempre fechar a conexão com o banco 
            finally
            {
                conexao.Close();
            }
        }

        private void txtId_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar)))
            {
                e.Handled = true;
            } 
        }

        

    }
}